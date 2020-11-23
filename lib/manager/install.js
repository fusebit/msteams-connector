/*
This is the installation logic of the Lifecycle Manager. 
*/

const Fs = require('fs');
const Sdk = require('./sdk'); // TODO @fusebit/add-on-sdk'

const getTemplateFiles = fileNames => fileNames.reduce((a, c) => {
    a[c] = Fs.readFileSync(__dirname + `/template/${c}`, { encoding: 'utf8' });
    return a;
}, {});

module.exports = async (ctx) => {
    const configuration = {
        ...ctx.body.configuration,
        ...Sdk.createStorage(ctx, true)
    };

    // Create the Add-On Handler
    await Sdk.createFunction(ctx, { 
        configurationSerialized: `# Add-on configuration settings
${Object.keys(configuration).sort().map(k => `${k}=${configuration[k]}`).join('\n')}
`,
        nodejs: {
            files: getTemplateFiles(['index.js', 'package.json', 'VendorBot.js']),
        },
        metadata: {
            fusebit: {
                editor: {
                    navigationPanel: {
                        hideFiles: [],
                    }
                }
            },
            ...ctx.body.metadata
        },
        functionPermissions: {
            allow: [
              {
                action: 'storage:*',
                resource: `/account/${ctx.accountId}/subscription/${ctx.subscriptionId}/storage/boundary/${ctx.body.boundaryId}/function/${ctx.body.functionId}/`
              },
              {
                action: 'function:*',
                resource: `/account/${ctx.accountId}/subscription/${ctx.subscriptionId}/`
              }
            ]
        },
        authentication: 'optional'
    });
    
    return { status: 200, body: { status: 200 }};
};

