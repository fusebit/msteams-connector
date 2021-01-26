const Sdk = require('@fusebit/add-on-sdk');
const { createApp } = require('./app');
const { FusebitBot } = require('./FusebitBot');

exports.FusebitBot = FusebitBot;

exports.createMicrosoftTeamsConnector = (vendorBot) => {
    // Create Express app that exposes endpoints to receive notifications from Teams, handle vendor authorization,
    // and sending of notifications. Teams notifications are handled by the vendor's bot.
    const app = createApp(vendorBot);

    // Create Fusebit function from the Express app
    const handler = Sdk.createFusebitFunctionFromExpress(app);

    return handler;
};
