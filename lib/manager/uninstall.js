/*
This is the uninstallation logic of the Lifecycle Manager. 
*/

const Sdk = require('@fusebit/add-on-sdk');
const Superagent = require('superagent');

module.exports = async (ctx) => {
    // Let the Connector clean up its internal state
    await Superagent.delete(
        `${ctx.body.baseUrl}/v1/account/${ctx.body.accountId}/subscription/${ctx.body.subscriptionId}/boundary/${ctx.body.boundaryId}/function/${ctx.body.functionId}`
    ).set('Authorization', `Bearer ${ctx.fusebit.callerAccessToken}`);

    // Destroy the Add-On Handler
    await Sdk.deleteFunction(ctx, ctx.fusebit.functionAccessToken);

    return { status: 204 };
};
