const { VendorBot } = require('./VendorBot');
const { createMicrosoftTeamsConnector } = require('@fusebit/msteams-connector');

module.exports = createMicrosoftTeamsConnector(new VendorBot());
