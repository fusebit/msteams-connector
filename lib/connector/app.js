const Sdk = require('@fusebit/add-on-sdk');
const { BotFrameworkAdapter } = require('botbuilder');
const Crypto = require('crypto');
const Superagent = require('superagent');

exports.createApp = (bot) => {
    const app = require('express')();

    const adapter = new BotFrameworkAdapter({
        appId: process.env.microsoft_app_id,
        appPassword: process.env.microsoft_app_password,
    });

    adapter.onTurnError = async (context, error) => {
        return bot.onTurnError(context, error);
    };

    // Called from the template manager to clean up all subordinate artifacts of this handler
    app.delete(
        '/',
        bot.authorize({
            action: 'function:delete',
            resourceFactory: (req) =>
                `/account/${req.fusebit.accountId}/subscription/${req.fusebit.subscriptionId}/boundary/${req.fusebit.boundaryId}/function/${req.fusebit.functionId}/`,
        }),
        async (req, res, next) => {
            // Clean up storage
            await req.fusebit.storage.delete(undefined, true);

            // Delete handlers
            const baseUrl = await bot._getBaseUrl(req.fusebit);
            while (true) {
                const response = await Superagent.get(
                    `${baseUrl}/v1/account/${req.fusebit.accountId}/subscription/${req.fusebit.subscriptionId}/function`
                )
                    .query(`search=tag.ownerId=${req.fusebit.boundaryId}/${req.fusebit.functionId}`)
                    .query({ count: 20 })
                    .set('Authorization', `Bearer ${req.fusebit.fusebit.functionAccessToken}`);
                if (!response.body || !response.body.items || response.body.items.length === 0) {
                    break;
                }
                await Promise.all(
                    response.body.items.map((f) => {
                        return Superagent.delete(
                            `${baseUrl}/v1/account/${req.fusebit.accountId}/subscription/${req.fusebit.subscriptionId}/boundary/${f.boundaryId}/function/${f.functionId}`
                        ).set('Authorization', `Bearer ${req.fusebit.fusebit.functionAccessToken}`);
                    })
                );
            }

            res.send(204);
        }
    );

    // Messages from individual user handlers to the connector intended to send a notification to the
    // Microsoft Teams user associated with a specific vendor user.
    app.post(
        '/api/notification/:vendorUserId',
        bot.authorize({
            action: 'function:execute',
            resourceFactory: (req) =>
                `/account/${req.fusebit.accountId}/subscription/${req.fusebit.subscriptionId}/boundary/${req.fusebit.boundaryId}/function/${req.fusebit.functionId}/operation/notification/${req.params.vendorUserId}/`,
        }),
        async (req, res) => {
            let vendor2teams = await req.fusebit.storage.get(bot._getStorageIdForVendorUser(req.params.vendorUserId));
            vendor2teams = vendor2teams ? vendor2teams.data : undefined;
            let userContext = vendor2teams ? await req.fusebit.storage.get(vendor2teams.storageId) : undefined;
            userContext = userContext ? userContext.data : undefined;
            if (!userContext) {
                return res.status(404).json({
                    status: 404,
                    statusCode: 404,
                    message: `Vendor user ${req.params.vendorUserId} is not associated with a Microsoft Teams user`,
                });
            }
            let response;
            try {
                response = await bot.onNotification(userContext, req.fusebit.body, req.fusebit);
            } catch (e) {
                const status = e.status || 500;
                return res.status(status).json({
                    status,
                    statusCode: status,
                    message: e.stack || e.messgage || e,
                });
            }
            return res.send(response);
        }
    );

    // Messages from Microsoft Teams
    app.post('/api/messages', async (req, res, next) => {
        req.body = req.fusebit.body;
        Sdk.debug('TEAMS REQUEST', req.body);
        try {
            await adapter.processActivity(req, res, async (context) => {
                context.fusebit = req.fusebit;
                Sdk.debug('PROCESSING TURN CONTEXT', context.activity);
                await bot.run(context);
            });
        } catch (e) {
            Sdk.debug('TEAMS REQUEST PROCESSING ERROR', e);
            next(e);
        }
    });

    // OAuth start page that redirects to vendor's authorization server.
    // Required because it must be on the same domain as callback URL.
    app.get('/start-oauth', async (req, res, next) => {
        res.send(await bot.getOAuthStartPageHtml(req.fusebit.configuration.vendor_oauth_authorization_url));
    });

    // OAuth callback from vendor's authorization server
    app.get('/callback', async (req, res, next) => {
        Sdk.debug('OAUTH CALLBACK', req.query);
        if (!req.query.state) {
            return res.send(await bot.getOAuthErrorPageHtml('The OAuth callback does not specify the `state` query parameter.'));
        }
        let state;
        try {
            state = JSON.parse(Buffer.from(req.query.state, 'hex'));
            if (typeof state.nonce !== 'string' || typeof state.user !== 'string') {
                throw new Error();
            }
        } catch (_) {
            return res.send(await bot.getOAuthErrorPageHtml('The `state` query parameter is malformed.'));
        }
        const storageKey = bot.getStorageIdForTeamsUser(state.user);
        let teamsUser;
        try {
            teamsUser = await req.fusebit.storage.get(storageKey);
            teamsUser = teamsUser ? teamsUser.data : undefined;
            await req.fusebit.storage.delete(storageKey);
            if (!teamsUser || teamsUser.state !== req.query.state || teamsUser.status !== 'authenticating') {
                throw new Error();
            }
        } catch (_) {
            return res.send(
                await bot.getOAuthErrorPageHtml('The authorization transaction has been tampered with or was restarted by the user.')
            );
        }
        if (req.query.error || !req.query.code) {
            return res.send(
                await bot.getOAuthErrorPageHtml(req.query.error || 'The OAuth callback does not specify the `code` query parameter.')
            );
        }
        const verificationCode = Crypto.randomBytes(8).toString('hex').substring(0, 4);
        try {
            const vendorToken = await bot.getAccessToken(req.query.code, `${req.fusebit.baseUrl}/callback`);
            if (!isNaN(vendorToken.expires_in)) {
                vendorToken.expires_at = Date.now() + +vendorToken.expires_in * 1000;
            }
            await req.fusebit.storage.put(
                {
                    data: {
                        status: 'validating',
                        verificationCode,
                        cardActivityId: teamsUser.cardActivityId,
                        vendorToken,
                        timestamp: Date.now(),
                    },
                },
                storageKey
            );
        } catch (e) {
            Sdk.debug('AUTHORIZATION CODE EXCHANGE ERROR', e);
            return res.send(await bot.getOAuthErrorPageHtml('Error exchanging the authorization code for an access token.'));
        }
        res.send(await bot.getOAuthCallbackPageHtml(verificationCode));
    });

    bot.onCreate(app);

    return app;
};
