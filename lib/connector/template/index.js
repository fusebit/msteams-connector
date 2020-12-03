const Superagent = require('superagent');

/**
 * @param ctx {FusebitContext}
 */
module.exports = async (ctx) => {
    // Pass through the request body to the connector's /api/notification/:vendorUserId API
    // to allow the VendorBot to send a notification to Microsoft Teams's user associated
    // with the vendor user
    const response = await Superagent.post(
        `${ctx.configuration.connector_url}/api/notification/${encodeURIComponent(ctx.configuration.vendor_user_id)}`
    )
        .set('Authorization', `Bearer ${ctx.fusebit.functionAccessToken}`)
        .send(ctx.body)
        .ok((res) => true);
    return { status: response.status, body: response.body };
};
