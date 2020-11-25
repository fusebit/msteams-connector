const { TurnContext } = require('botbuilder');
const { FusebitBot } = require('@fusebit/msteams-connector');

class VendorBot extends FusebitBot {
    constructor() {
        super();
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);
            const text = context.activity.text.trim().toLocaleLowerCase();

            if (text.includes('login')) {
                await this.sendSignInCardAsync(context);
            }
            else if (text.includes('logout')) {
                await this.deleteTeamsUser(context);
                await context.sendActivity("You are logged out. Use 'login' command to log in.");
            }
            else {
                const userContext = await this.getUserContext(context);
                if (userContext) {
                    await context.sendActivity(`Welcome! Your login status is '${userContext.status}'. Use 'logout' to log out.`);
                }
                else {
                    await context.sendActivity(`Welcome! Use 'login' command to log in.`);
                }
            }

            await next();
        });
    }

    /**
     * Invoked when a call from a vendor's user handler requested a notification to be sent to a Teams user
     * associated with a specific vendor's user. The return value will be JSON-serialized and passed back to the 
     * vendor's user handler.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively.
     * @param {*} payload The payload is an object passed-through unchanged from the request from the vendor's user hadler. 
     * @param {FusebitContext} fusebitContext The FusebitContext representing the request
     */
    async onNotification(userContext, payload, fusebitContext) {
        // Default implementation returns an HTTP 501 Not Implemented
        // Implement this method to enable sending notifications from the vendor's system to Microsoft Teams.
        // Documentation: https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/send-proactive-messages?tabs=typescript
        return super.onNotification(userContext, payload, fusebitContext);;
    }


    /**
     * Called before creating a Fusebit Function responsible for handling logic specific to a user described by the userContext. 
     * This is an opportunity to modify the function specification, for example to:
     * - change the files making up the default implementation,
     * - add tags to allow for criteria-based lookup of specific functions using Fusebit APIs,
     * - add custom configuration settings.
     * The new function specification must be returned as a return value. 
     * @param {TurnContext} context The TurnContext of the Microsoft Bot Framework.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively.
     * @param {*} functionSpecification Default Fusebit Function specification
     */
    async modifyFunctionSpecification(context, userContext, functionSpecification) {
        return functionSpecification;
    }

    /**
     * Return Fusebit Function boundary Id for the specified user context.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively. 
     */
    async getBoundaryId(userContext) {
        // The base class returns 'msteams-user-{first-40-characters-of-hex-encoded-sha1-of-teams-user-id}'
        return super.getBoundaryId(userContext);
    }

    /**
     * Return Fusebit Function function Id for the specified user context.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively. 
     */
    async getFunctionId(userContext) {
        // The base class returns 'msteams-handler'
        return super.getFunctionId(userContext);
    }

    /**
     * Called when a Microsoft Teams user successfuly authorized access to the vendor's system.
     * @param {TurnContext} context The TurnContext of the Microsoft Bot Framework.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively.
     */
    async onUserLoggedIn(context, userContext) {
        await context.sendActivity('Thank you for logging in!');
    };

    /**
     * Creates the fully formed web authorization URL to vendor's system to start the authorization flow.
     * @param {string} state The value of the OAuth state parameter.
     * @param {string} redirectUri The callback URL to redirect to after the authorization flow.
     */
    async getAuthorizationUrl(state, redirectUri) {
        // The base class implementation constructs the authorization URL using the OAuth parameters stored
        // in the function's configuration.  
        return super.getAuthorizationUrl(state, redirectUri);
    };

    /**
     * Exchanges the OAuth authorization code for the access and refresh tokens. Returns a JSON object with all OAuth token 
     * exchange parameters (access_token, expires_in, refresh_token etc). The JSON object is called token context.
     * @param {string} authorizationCode The authorization_code supplied to the OAuth callback upon successful authorization flow.
     * @param {string} redirectUri The redirect_uri value Fusebit used to start the authorization flow. 
     */
    async getAccessToken(authorizationCode, redirectUri) {
        // The base class implementation exchanges the authorization code for an access token
        // using the OAuth parameters stored in the function's confguration.  
        return super.getAccessToken(authorizationCode, redirectUri);
    };

    /**
     * Obtains a new access token using refresh token.
     * @param {*} tokenContext An object representing the result of the getAccessToken call. It contains refresh_token.
     */
    async refreshAccessToken(tokenContext) {
        // The base class implementation issues a POST request with OAuth parameters in the query string
        // using the OAuth parameters stored in the function's confguration.  
        return super.refreshAccessToken(tokenContext);
    };

    /**
     * Obtains the user profile after a newly completed authorization flow. User profile will be stored along the token
     * context and associated with Microsoft Teams user, and can be later used to customize the conversation with the Microsoft
     * Teams user.
     * @param {*} tokenContext An object representing the result of the getAccessToken call. It contains access_token.
     */
    async getUserProfile(tokenContext) {
        // Example of obtaining a user profile:
        // const response = await require('superagent').get('https://contoso.com/v1/users/me')
        //     .set('Authorization', `Bearer ${tokenContext.access_token}`);
        // return response.body;

        // The base class returns an empty object
        return super.getUserProfile();
    };

   /**
     * Returns a string uniquely identifying the user in vendor's system. Typically this is a property of 
     * userContext.vendorUserProfile. Default implementation is opportunistically returning userContext.vendorUserProfile.id
     * if it exists.
     * @param {*} userContext The user context representing the vendor's user. Contains vendorToken and vendorUserProfile, representing responses
     * from getAccessToken and getUserProfile, respectively.
     */
    async getUserId(userContext) {
        return super.getUserId(userContext);
    };

    /**
     * Called when the bot error occurred. 
     * @param {TurnContext} context The turn context 
     * @param {*} error The error
     */
    async onTurnError(context, error) {
        return super.onTurnError(context, error);
    };

    /**
     * Generates the HTML of the web page that is returned from the OAuth callback endpoint upon 
     * successful authorization flow against vendor's system. The page must call 
     * microsoftTeams.authentication.notifySuccess method with the specified verificationCode, 
     * as documented at https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-flow-bot
     * @param {string} verificationCode A one-time verification code that must be supplied back to Microsoft Teams
     */
    async getOAuthCallbackPageHtml(verificationCode) {
        return super.getOAuthCallbackPageHtml(verificationCode);
    };

    /**
     * Generates the HTML of the web page that is returned from the OAuth callback endpoint upon 
     * authorization error from the vendor's system. 
     * @param {string} reason A descriptive reason for the authorization failure
     */
    async getOAuthErrorPageHtml(reason) {
        return super.getOAuthErrorPageHtml(reason);
    };

    /**
     * Generates the HTML of the web page that must redirect the browser to the URL supplied through the 
     * 'authorizationUrl' query parameter. This is required because the OAuth authorization flow in Microsoft
     * Teams must originate on the same domain name as the authorization callback, as described in
     * https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-flow-bot.
     * The page should validate that the 'authorizationUrl' query parameter has the same base URL as the 
     * authorizationUrlBase before executing the redirect, in order to mitigate misuse of the endpoint. 
     * @param {string} authorizationUrlBase The expected base URL of the 'authorizationUrl' query parameter to enforce
     */
    async getOAuthStartPageHtml(authorizationUrlBase) {
        return super.getOAuthStartPageHtml(authorizationUrlBase);
    };

}

module.exports.VendorBot = VendorBot;
