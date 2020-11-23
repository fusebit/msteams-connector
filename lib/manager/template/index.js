const { VendorBot } = require('./VendorBot');
const { createMicrosoftTeamsAddon } = require('@fusebit/msteams-add-on');

module.exports = createMicrosoftTeamsAddon(new VendorBot());
