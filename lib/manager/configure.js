/*
This is the state machine that describes the configuration logic of the Lifecycle Manager. 
*/

const Sdk = require('@fusebit/add-on-sdk');
const form = require('fs').readFileSync(__dirname + '/form.html', {
    encoding: 'utf8',
});

module.exports = {
    initialState: 'settingsManagers',
    states: {
        settingsManagers: async (ctx, state, data) => {
            const settingsManagers = [];
            (ctx.configuration.fusebit_settings_managers || '').split(',').forEach((s) => {
                if (s.trim()) {
                    settingsManagers.push(s);
                }
            });
            const stage = state.settingsManagersStage || 0;
            if (settingsManagers.length > stage) {
                // Invoke subsequent settings manager
                state.settingsManagersStage = stage + 1;
                return Sdk.redirect(ctx, state, data, settingsManagers[stage], 'settingsManagers');
            } else {
                // All settings managers processed (or none defined), move to the 'form' state
                state.configurationState = 'form';
                return await module.exports.states.form(ctx, state, data);
            }
        },
        form: async (ctx, state, data) => {
            if (!ctx.configuration.fusebit_show_form_configuration) {
                // Do not show web form configuration, complete the configuration stage
                return Sdk.completeWithSuccess(state, data);
            }

            // Render a simple HTML form to collect configuration parameters required by the Add-On.
            // The form will perform a client-side redirect back to the `ctx.query.returnTo`, which is
            // the application that initiated the configuration flow. The application will normally
            // continue the Add-On installation by invoking the /install endpoint on the Lifecycle Manager.

            const view = form
                .replace(/##templateName##/g, data.templateName)
                .replace(/##returnTo##/, JSON.stringify(state.returnTo))
                .replace(/##data##/, JSON.stringify(data))
                .replace(/##state##/, state.returnToState ? JSON.stringify(state.returnToState) : 'null');

            return {
                body: view,
                bodyEncoding: 'utf8',
                headers: { 'content-type': 'text/html' },
                status: 200,
            };
        },
    },
};
