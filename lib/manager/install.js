/*
This is the installation logic of the Lifecycle Manager. 
*/

const Fs = require('fs');
const Sdk = require('@fusebit/add-on-sdk');

const getTemplateFiles = (fileNames) =>
    fileNames.reduce((a, c) => {
        a[c] = Fs.readFileSync(__dirname + `/template/${c}`, { encoding: 'utf8' });
        return a;
    }, {});

module.exports = async (ctx) => {
    const configuration = {
        debug: '1',
        ...ctx.body.configuration,
    };

    // Create the Connector
    await Sdk.createFunction(
        ctx,
        {
            configurationSerialized: `# Connector configuration settings
${Object.keys(configuration)
    .sort()
    .map((k) => `${k}=${configuration[k]}`)
    .join('\n')}
`,
            nodejs: {
                files: {
                    ...getTemplateFiles(['index.js', 'VendorBot.js']),
                    'package.json': {
                        engines: {
                            node: '10',
                        },
                        dependencies: {
                            // Use the same version of @fusebit-msteams-connector as in this package
                            '@fusebit/msteams-connector': require('../../package.json').version,
                            // The following two are declared as peerDependencies of @fusebit/msteams-connector
                            // and the versions must match those in top level package.json
                            botbuilder: '4.11.0',
                            superagent: '6.1.0',
                        },
                    },
                },
            },
            metadata: {
                fusebit: {
                    editor: {
                        navigationPanel: {
                            hideFiles: [],
                        },
                    },
                },
                ...ctx.body.metadata,
            },
            security: {
                functionPermissions: {
                    allow: [
                        {
                            action: 'storage:*',
                            resource: `/account/${ctx.accountId}/subscription/${ctx.subscriptionId}/storage/boundary/${ctx.body.boundaryId}/function/${ctx.body.functionId}/`,
                        },
                        {
                            action: 'function:*',
                            resource: `/account/${ctx.accountId}/subscription/${ctx.subscriptionId}/`,
                        },
                    ],
                },
                authentication: 'optional',
            },
        },
        ctx.fusebit.functionAccessToken
    );

    return { status: 200, body: { status: 200 } };
};
