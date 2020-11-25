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

* `lib/connector` contains the core Fusebit Microsoft Team Connector logic that implements the two-way integration between your SaaS and Microsoft Teams.
* `lib/manager` contains the Fusebit Microsoft Team Connector Manager logic which supports the install/uninstall/configure operations for the connector.
* `lib/manager/template` contains a template a Fusebit Function that exposes the Fusebit Microsoft Team Connector interface. As a developer, you will be spending most of your time focusing on adding your integration logic to [VendorBot.js](https://github.com/fusebit/msteams-connector/blob/main/lib/manager/template/VendorBot.js). 
* `fusebit` contains a template of a Fusebit Function that exposes the Fusebit Microsoft Team Connector Manager interface. 

## Release notes

### v2.0.0

* Removed FusebitBot.getBoundaryIdForTeamsUser and Fusebit.getFunctionIfForTeamsUser. 
* Added FusebitBot.getBoundaryId and FusebitBot.getFunctionId that accept the entire userContext instead. 

### v1.0.0

* Initial implementation.