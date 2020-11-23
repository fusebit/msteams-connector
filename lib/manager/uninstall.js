/*
This is the uninstallation logic of the Lifecycle Manager. 
*/

const Sdk = require('@fusebit/add-on-sdk');

module.exports = async (ctx) => {
    // Destroy the Add-On Handler
    await Sdk.deleteFunction(ctx);

    return { status: 204 };
};
