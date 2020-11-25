const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
} = require('botbuilder');
const Crypto = require('crypto');
const Sdk = require('../sdk'); // TODO @fusebit/add-on-sdk
const html = require('./html');
const Superagent = require('superagent');
const Fs = require('fs');

const getTemplateFile = (fileName) => Fs.readFileSync(__dirname + `/template/${fileName}`, { encoding: 'utf8' });

class FusebitBot extends TeamsActivityHandler {
    constructor() {
        super();
    }

    /**
     * Called before creating a Fusebit Function responsible for handling logic specific to a user described by the userContext. 
     * This is an opportunity to modify the function specification, for example to:
     * - change the files making up the default implementation of the function,
     * - add tags to allow for criteria-based lookup of specific handlers using Fusebit APIs,
     * - add custom configuration settings.
     * The new function specification must be returned as a return value. 
     * @param {TurnContext} context The TurnContext of the Microsoft Bot Framework.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively.
     * @param {*} functionSpecification Default Fusebit Function specification
     */
    async modifyFunctionSpecification(context, userContext, functionSpecification) {
        return functionSpecification;
    }

    /**
     * Invoked when a call from a vendor's user handler requested a notification to be sent to a Teams user
     * associated with a specific vendor's user. The return value will be JSON-serialized and passed back to the 
     * vendor's user handler.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively.
     * @param {*} payload The payload is an object passed-through unchanged from the request from the vendor's user hadler. 
     * @param {FusebitContext} fusebitContext The FusebitContext representing the request
     */
    async onNotification(userContext, payload, fusebitContext) {
        const error = new Error('Not implemented. Please implement VendorBot.onNotification to process notificiations to teams users.');
        error.status = error.statusCode = 501;
        throw error;
    }

    /**
     * Called when a Microsoft Teams user successfuly authorized access to the vendor's system.
     * @param {TurnContext} context The TurnContext of the Microsoft Bot Framework.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively.
     */
    async onUserLoggedIn(context, userContext) {
    };

    /**
     * Creates the fully formed web authorization URL to start the authorization flow.
     * @param {string} state The value of the OAuth state parameter.
     * @param {string} redirectUri The callback URL to redirect to after the authorization flow.
     */
    async getAuthorizationUrl(state, redirectUri) {
        return [
            process.env.vendor_oauth_authorization_url,
            `?response_type=code`,
            `&scope=${encodeURIComponent(process.env.vendor_oauth_scope)}`,
            `&state=${state}`,
            `&client_id=${process.env.vendor_oauth_client_id}`,
            `&redirect_uri=${encodeURIComponent(redirectUri)}`
        ].join('');
    };

    /**
     * Exchanges the OAuth authorization code for the access and refresh tokens.
     * @param {string} authorizationCode The authorization_code supplied to the OAuth callback upon successful authorization flow.
     * @param {string} redirectUri The redirect_uri value Fusebit used to start the authorization flow. 
     */
    async getAccessToken(authorizationCode, redirectUri) {
        const response = await Superagent.post(process.env.vendor_oauth_token_url)
            .type('form')
            .send({
                grant_type: 'authorization_code',
                code: authorizationCode,
                client_id: process.env.vendor_oauth_client_id,
                client_secret: process.env.vendor_oauth_client_secret,
                redirect_uri: redirectUri,
            });
        return response.body;
    };

    /**
     * Obtains a new access token using refresh token.
     * @param {*} tokenContext An object representing the result of the getAccessToken call. It contains refresh_token.
     */
    async refreshAccessToken(tokenContext) {
        const response = await Superagent.post(process.env.vendor_oauth_token_url)
            .query({
                grant_type: 'refresh_token',
                refresh_token: tokenContext.refresh_token,
                client_id: process.env.vendor_oauth_client_id,
                client_secret: process.env.vendor_oauth_client_secret,
            });
        return response.body;
    };

    /**
     * Obtains the user profile given a freshly completed authorization flow. User profile will be stored along the token
     * context and associated with Microsoft Teams user, and can be later used to customize the conversation with the Microsoft
     * Teams user.
     * @param {*} tokenContext An object representing the result of the getAccessToken call. It contains access_token.
     */
    async getUserProfile(tokenContext) {
        return {};
    };

    /**
     * Returns a string uniquely identifying the user in vendor's system. Typically this is a property of 
     * userContext.vendorUserProfile. Default implementation is opportunistically returning userContext.vendorUserProfile.id
     * if it exists.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively.
     */
    async getUserId(userContext) {
        if (userContext.vendorUserProfile.id) {
            return userContext.vendorUserProfile.id;
        }
        throw new Error('Please implement the getUserId method in the class deriving from FusebitBot.');
    };

    /**
     * Called when the bot error occurred. 
     * @param {TurnContext} context The turn context 
     * @param {*} error The error
     */
    async onTurnError(context, error) {
        Sdk.debug(`TURN ERROR: ${error.stack || error}`);
    
        // Send a trace activity, which will be displayed in Bot Framework Emulator
        await context.sendTraceActivity(
            'OnTurnError Trace',
            `${error}`,
            'https://www.botframework.com/schemas/error',
            'TurnError'
        );
    
        // Send a message to the user
        await context.sendActivity(`The bot encountered an error: ${error}`);
    };

    /**
     * Generates the HTML of the web page that is returned from the OAuth callback endpoint upon 
     * successful authorization flow against vendor's system. The page must call 
     * microsoftTeams.authentication.notifySuccess method with the specified verificationCode, 
     * as documented at https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-flow-bot
     * @param {string} verificationCode A one-time verification code that must be supplied back to Microsoft Teams
     */
    async getOAuthCallbackPageHtml(verificationCode) {
        return html.getOAuthCallbackPageHtml(verificationCode);
    };

    /**
     * Generates the HTML of the web page that is returned from the OAuth callback endpoint upon 
     * authorization error from the vendor's system. 
     * @param {string} reason A descriptive reason for the authorization failure
     */
    async getOAuthErrorPageHtml(reason) {
        return html.getOAuthErrorPageHtml(reason);
    };

    /**
     * Generates the HTML of the web page that must redirect the browser to the URL supplied through the 
     * 'authorizationUrl' query parameter. This is required because the OAuth authorization flow in Microsoft
     * Teams must originate on the same domain name as the authorization callback, as described in
     * https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-flow-bot.
     * The page should validate that the 'authorizationUrl' query parameter has the same base URL as the 
     * authorizationUrlBase before executing the redirect, in order to mitigate misuse of the endpoint. 
     * @param {string} authorizationUrlBase The expected base URL of the 'authorizationUrl' query parameter to enforce
     */
    async getOAuthStartPageHtml(authorizationUrlBase) {
        return html.getOAuthStartPageHtml(authorizationUrlBase);
    };

    /**
     * Gets the user context representing the association of a Microsoft Teams user with a vendor's user. Returns an object
     * that contains vendorToken and vendorUserProfile, representing responses getAccessToken and getUserProfile, respectively, 
     * as well as a teamsUser object representing the Teams user, channel, team, and tenant. 
     * If the user who sent the turn context is not logged in to the vendor's system, undefined is returned. 
     * @param {TurnContext} context The turn context
     */
    async getUserContext(context) {
        return context.fusebit.storage.get(this._getStorageIdForTeamsUser(context.activity.from.id));
    }
    
    /**
     * Save an updated user context in storage for future use.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively.
     */
    async saveUserContext(userContext) {
        return context.fusebit.storage.put(userContext, true, this._getStorageIdForTeamsUser(userContext.teamsUser.user));
    }

    /**
     * Returns a valid access token to the vendor's system representing the vendor's user described by the userContext.
     * If the currently stored access token is expired or nearing expiry, and a refresh token is available, a new access
     * token is obtained, stored for future use, and returned. If a current access token cannot be returned, an exception is thrown.
     * @param {TurnContext} context The turn context 
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively. This is optional and if not provided, will be obtained from context.
     */
    async ensureAccessToken(context, userContext) {
        if (!userContext) {
            userContext = await this.getUserContext(context);
        }
        if (!userContext) {
            throw new Error(`Cannot return an access token because the user is not authenticated.`);
        }
        if (userContext.status !== 'authenticated') {
            throw new Error(`Cannot return an access token because the user is not authenticated. Current status is '${userContext.status}'`);
        }
        if (userContext.vendorToken) {
            if (userContext.vendorToken.access_token && userContext.vendorToken.expires_at > (Date.now() + 30000)) {
                return userContext.vendorToken.access_token;
            }
            if (userContext.vendorToken.refresh_token) {
                userContext.status = 'refreshing';
                try {
                    await this.saveUserContext(userContext);
                    userContext.vendorToken = await this.refreshAccessToken(userContext.vendorToken);
                    if (!isNaN(userContext.vendorToken.expires_in)) {
                        userContext.vendorToken.expires_at = Date.now() + +userContext.vendorToken.expires_in * 1000;
                    }        
                    userContext.vendorUserProfile = await this.getUserProfile(userContext.vendorToken);
                    userContext.status = 'authenticated';
                    await this.saveUserContext(userContext);
                    return userContext.vendorToken.access_token;
                }
                catch (e) {
                    await this.deleteVendorUser(context.fusebit, userContext);
                    Sdk.debug('REFRESH TOKEN ERROR', e);
                }
            }
        }
        throw new Error('User logged out. Unable to obtain an access token.');
    }

    async handleTeamsSigninVerifyState(context, query) {
        Sdk.debug('VERIFY STATE', query);
        const storageKey = this._getStorageIdForTeamsUser(context.activity.from.id);
        let teamsUser;
        try {
            teamsUser = await context.fusebit.storage.get(storageKey);
            await context.fusebit.storage.delete(false, storageKey);
            if (teamsUser.status !== 'validating' || teamsUser.verificationCode !== query.state) {
                throw new Error();
            }
        }
        catch (_) {
            await context.sendActivity('You are not logged in. Integrity of the authentication transaction could not be validated.');
            return;
        }

        const userContext = {
            status: 'authenticated',
            vendorToken: teamsUser.vendorToken,
            vendorUserProfile: await this.getUserProfile(teamsUser.vendorToken),
            teamsUser: {
                user: context.activity.from.id,
                channel: context.activity.channelData.channel.id,
                team: context.activity.channelData.team.id,
                tenant: context.activity.channelData.tenant.id,
            }
        };
        userContext.vendorUserId = await this.getUserId(userContext);

        const ctx = this._getSyntheticFusebitFunctionContextFromTurnContext(context);

        // Function specification of the handler specific to the vendor's user
        let functionSpecification = {
            nodejs: {
                files: {
                    'package.json': {
                        engines: {
                            node: '10',
                        },
                        dependencies: {
                            'superagent': '6.1.0'
                        },
                    },
                    'index.js': getTemplateFile('index.js'),
                },
            },
            metadata: {
                tags: {
                    'msteamsUser': userContext.teamsUser.user,
                    'msteamsChannel': userContext.teamsUser.channel,
                    'msteamsTeam': userContext.teamsUser.team,
                    'msteamsTenant': userContext.teamsUser.tenant,
                    'vendorUser': userContext.vendorUserId,
                    'ownerId': `${context.fusebit.boundaryId}/${context.fusebit.functionId}`
                },
                fusebit: {
                    editor: {
                        navigationPanel: {
                            hideFiles: [],
                        },
                    },
                },
            },
            security: {
                // Permit the handler to call back to the connector to send a notification message to Teams
                // on behalf of a specific vendor's user
                functionPermissions: {
                    allow: [
                        {
                            action: 'function:execute',
                            resource: `/account/${context.fusebit.accountId}/subscription/${context.fusebit.subscriptionId}/boundary/${context.fusebit.boundaryId}/function/${context.fusebit.functionId}/operation/notification/${userContext.vendorUserId}/`
                        }
                    ]
                },
                // All callers to handler MUST have function:execute to the handler
                authentication: 'required',
                authorization: [
                    {
                        action: 'function:execute',
                        resource: `/account/${context.fusebit.accountId}/subscription/${context.fusebit.subscriptionId}/boundary/${ctx.body.boundaryId}/function/${ctx.body.functionId}/`
                    }
                ],
            },
            configurationSerialized: `# Vendor's user ID
vendor_user_id=${userContext.vendorUserId}

# Connector URL
connector_url=${context.fusebit.baseUrl}
`
        };

        functionSpecification = await this.modifyFunctionSpecification(context, userContext, functionSpecification);
        userContext.functionUrl = await Sdk.createFunction(ctx, functionSpecification, context.fusebit.fusebit.functionAccessToken);
        let storageCreated;
        try {
            await context.fusebit.storage.put(userContext, true, storageKey);
            storageCreated = true;
            await context.fusebit.storage.put({ storageId: storageKey }, true, this._getStorageIdForVendorUser(userContext.vendorUserId));
        }
        catch (e) {
            if (storageCreated) {
                await context.fusebit.storage.delete(false, storageKey);
            }
            await Sdk.deleteFunction(ctx, undefined, undefined, context.fusebit.fusebit.functionAccessToken);
            throw e;
        }
        await this.onUserLoggedIn(context, userContext);
    }

    /**
     * Sends a signin card to the Microsoft Teams user which allows the user to log into the vendor's system
     * to establish an association with the the vendor's user. 
     * @param {TurnContext} context The turn context
     */
    async sendSignInCardAsync(context) {
        const state = Buffer.from(JSON.stringify({
            nonce: Crypto.randomBytes(32).toString('base64'),
            user: context.activity.from.id
        })).toString('hex');
        await context.fusebit.storage.put({ status: 'authenticating', state, timestamp: Date.now() }, true, this._getStorageIdForTeamsUser(context.activity.from.id));
        const authorizationUrl = await this.getAuthorizationUrl(state, `${context.fusebit.baseUrl}/callback`);
        Sdk.debug('AUTHORIZATION URL', authorizationUrl);
        const startAuthorizationUrl = [
            context.fusebit.baseUrl,
            '/start-oauth',
            `?authorizationUrl=${encodeURIComponent(authorizationUrl)}`
        ].join('');
        const card = CardFactory.signinCard(
            'Sign in',
            startAuthorizationUrl,
            `Welcome to ${context.fusebit.configuration.vendor_name}`
        );
        await context.sendActivity(MessageFactory.attachment(card));
    }

    _getStorageIdForTeamsUser(id) {
        return `teams-user/${Buffer.from(id).toString('hex')}`;
    }

    _getStorageIdForVendorUser(id) {
        return `vendor-user/${Buffer.from(id).toString('hex')}`;
    }

    /**
     * Return Fusebit Function boundary Id for the specified Microsoft Teams user Id.
     * @param {string} id Microsoft Teams user Id
     */
    async getBoundaryIdForTeamsUser(id) {
        return `msteams-user-${Crypto.createHash('sha1').update(id).digest('hex').substring(0, 40)}`;
    }

    /**
     * Return Fusebit Function function Id for the specified Microsoft Teams user Id.
     * @param {string} id Microsoft Teams user Id
     */
    async getFunctionIdForTeamsUser(id) {
        return `msteams-handler`;
    }

    /**
     * Creates a synthetic Fusebit ctx from the information in the turn context, to be used with Sdk.createFunction and
     * Sdk.deleteFunction.
     * @param {*} context The turn context
     */
    _getSyntheticFusebitFunctionContextFromTurnContext(context) {
        return { body: {
            baseUrl: context.fusebit.configuration.fusebit_storage_audience,
            accountId: context.fusebit.accountId, 
            subscriptionId: context.fusebit.subscriptionId,
            boundaryId: await this.getBoundaryIdForTeamsUser(context.activity.from.id),
            functionId: await this.getFunctionIdForTeamsUser(context.activity.from.id),
        }};
    }

    /**
     * Creates a synthetic Fusebit ctx from the information in the turn context, to be used with Sdk.createFunction and
     * Sdk.deleteFunction.
     * @param {*} fusebitContext The Fusebit context of the current request.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively. 
     */
    _getSyntheticFusebitFunctionContextFromUserContext(fusebitContext, userContext) {
        return { body: {
            baseUrl: fusebitContext.configuration.fusebit_storage_audience,
            accountId: fusebitContext.accountId, 
            subscriptionId: fusebitContext.subscriptionId,
            boundaryId: await this.getBoundaryIdForTeamsUser(userContext.teamsUser.user),
            functionId: await this.getFunctionIdForTeamsUser(userContext.teamsUser.user),
        }};
    }

    /**
     * Removes all artifacts associated with a Teams user. Should be called when the Teams user logs out from the system.
     * @param {TurnContext} context The turn context 
     */
    async deleteTeamsUser(context, userContext) {
        if (!userContext) {
            userContext = await this.getUserContext(context);
        }
        if (userContext) {
            await Sdk.deleteFunction(
                this._getSyntheticFusebitFunctionContextFromTurnContext(context), 
                undefined, 
                undefined, 
                context.fusebit.fusebit.functionAccessToken
            );
            await context.fusebit.storage.delete(false, this._getStorageIdForVendorUser(userContext.vendorUserId));
            await context.fusebit.storage.delete(false, this._getStorageIdForTeamsUser(context.activity.from.id));
        }
    }

    /**
     * Removes all artifacts associated with a vendor user. Should be called when the vendor user chooses to disassociate
     * from Teams.
     * @param {*} fusebitContext The Fusebit context of the current request.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively. 
     */
    async deleteVendorUser(fusebitContext, userContext) {
        await Sdk.deleteFunction(
            this._getSyntheticFusebitFunctionContextFromUserContext(fusebitContext, userContext), 
            undefined, 
            undefined, 
            fusebitContext.fusebit.functionAccessToken
        );
    }
}

exports.FusebitBot = FusebitBot;
