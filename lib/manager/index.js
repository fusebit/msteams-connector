/*
This is the entry point of the Lifecycle Manager.

It presents a simple HTML form to collect the configuration paramaters for the Add-On. 
*/

const configure = require('./configure');
const install = require('./install');
const uninstall = require('./uninstall');

const Sdk = require('@fusebit/add-on-sdk');

module.exports = Sdk.createLifecycleManager({ configure, install, uninstall });
