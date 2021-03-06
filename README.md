# Fusebit Connector for Microsoft Teams

This is the Fusebit Microsoft Teams Connector, a simple way to to implement two-way integration between your multi-tenant SaaS application and Microsoft Teams, on top of the [Fusebit](https://fusebit.io) platform.

## Getting started

Assuming you are a subscriber of [Fusebit](https://fusebit.io), you would start by using the `fuse` CLI to deploy a Fusebit Microsoft Team Connector Manager to your subscription:

```
git clone git@github.com:fusebit/msteams-connector.git
cd msteams-connector
fuse function deploy --boundary managers msteams-connector-manager -d ./fusebit
```

Soon enough you will be writing code of your integration logic. Get in touch at [Fusebit](https://fusebit.io) for further instructions or to learn more.

## Organization

-   `lib/connector` contains the core Fusebit Microsoft Team Connector logic that implements the two-way integration between your SaaS and Microsoft Teams.
-   `lib/manager` contains the Fusebit Microsoft Team Connector Manager logic which supports the install/uninstall/configure operations for the connector.
-   `lib/manager/template` contains a template a Fusebit Function that exposes the Fusebit Microsoft Team Connector interface. As a developer, you will be spending most of your time focusing on adding your integration logic to [VendorBot.js](https://github.com/fusebit/msteams-connector/blob/main/lib/manager/template/VendorBot.js).
-   `fusebit` contains a template of a Fusebit Function that exposes the Fusebit Microsoft Team Connector Manager interface.

## Release notes

### v2.4.1

-   Add sign-in card activity ID to user context when validating the token.

### v2.3.0

-   Add sign-in card activity ID to user context when sending a sign-in card.
-   Add public FusebitBot.getStorageIdForTeamsUser and FusebitBot.getStorageIdForVendorUser methods.

### v2.2.0

-   Add FusebitBot.onCreate to allow for registration of custom Express routes on the bot.
-   Add FusebitBot.authorize to help with creation of authorization middleware to secure custom Express routes using Fusebit tokens.
-   Update @fusebit/add-on-sdk to 3.1.0 to simplify creation of Fusebit functions from Express router.

### v2.1.0

-   Fix bug to correctly pass the `payload` parameter to the FusebitBot.onNotification function.
-   Populate `userContext.teamsUser.conversation` and `userContext.teamsUser.serviceUrl` on completion of the vendor login flow.
-   Fix bug to properly handle vendor login completion initiated from a personal conversation with the bot. In those cases the `userContext.teamsUser.channel` and `userContext.teamsUser.team` will not be populated.
-   Replace embedded version of add-on-sdk with dependency on @fusebit/add-on-sdk.
-   Prettify everything.

### v2.0.1

-   Declaring botbuilder and superagent as peer dependencies to reduce the size of the deployed connector and its build time.

### v2.0.0

-   Removed FusebitBot.getBoundaryIdForTeamsUser and Fusebit.getFunctionIfForTeamsUser.
-   Added FusebitBot.getBoundaryId and FusebitBot.getFunctionId that accept the entire userContext instead.

### v1.0.0

-   Initial implementation.
