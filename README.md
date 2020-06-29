## K2 Microsoft Teams Connector

This connector for K2 Cloud integrates with Microsoft Teams. It allows you to create, update, and delete Teams, Channels in Teams, and Tabs in Channels. It also allows you to send a message to a channel.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

# Features

  - Work with Teams, Channels, and Tabs using SmartObjects
  - Send a message to a Channel
  - SVG icons if you want to add Custom Workflow Steps to K2 Designer for four methods: Create Team, Update Team, Create Channel, Send Message to Channel
  - Sample unit tests with mocks and code coverage
  - RollupJS configuration for TypeScript


## Getting the Microsoft Teams Connector to work in your K2 environment

Before you see SmartObjects and run methods on them, you must do the following:

 - Create an Azure App in your Azure tenant
 - Assign permissions (Application and/or Delegated) to the app so it has the necessary permissions
 in Microsoft Graph to work from K2. For more information about the rights required for each method, see the Help topic at https://k2.com/help/servicetypes/jssp/teams
 - Create a secret for your app and copy the secret and the ID of the app for configuring an OAuth resource in K2
 - Configure an OAuth Resource in K2 Management based on the Microsoft Online Resource Type
 - Create a Microsoft Teams service type with your bundled .js file (or use the index.jssp file in this repo's release)
 - Create a service instance for the type using the OAuth resource you created

 ### Creating the Azure App and Adding Permissions
 
 Create a new app in Azure by browsing to **Azure Active Directory > App registrations**. Give it a name like **K2 for Microsoft Teams**. Once you have the app created, make a note of its Application ID. This is the app's client ID that you'll use when you create the OAuth resource in K2.

 Also, you need to create a secret. Click Certificates & secrets and then click **New client secret**. Once you create the secret, make a note of it as you'll also need to add this to your OAuth Resource.

 Next, configure permissions by clicking **API permissions**. Permissions allow K2 to act on behalf of users. There are **Application** and **Delegated** permissions available in Azure. It's safest to grant both types to the app where specified. Click the **Add a permission** button and find the following Microsoft Graph permissions to add (search usually works best once you click Microsoft Graph):

  + K2 API (search for this one and assign Delegated permissions)
  + Directory.Read.All
  + email
  + Group.Create
  + Group.ReadWrite.All
  + openid
  + profile
  + TeamsActivity.Read.All
  + TeamsActivity.Send
  + TeamsApp.ReadWrite.All
  + TeamSettings.ReadWrite.All
  + TeamsTab.ReadWrite.All
  + User.Read

  Once you have all of these permissions granted, you must click the **Grant admin consent** button at the top of the page. This gives your new app the ability to perform Teams-related tasks suppported by the connector's methods. If you do not want to grant all of these permissions, you can grant a subset of the permissions but the connector methods will not function. If you choose, you can edit the connector source file directly to only have the methods you need for your Teams automation needs.

  Lastly, make a note of the endpoints in your environment. At the top of the **Overview** page for your app, click the **Endpoints** button. Here you need to copy the **OAuth 2.0 authorization endpoint (v2)** and the **OAuth 2.0 token endpoint (v2)**.  

  ### Creating the OAuth Resource in K2

  Next, create the OAuth resource in K2. Follow these steps:
  
  1. Browse to K2 Management and expand the **Authentication** node
  2. Expand the **OAuth** node
  3. Click **Resources** and then click **New**
  4. Give your resource a name, select **Microsoft Online** as the Resource Type, and then paste the authorization endpoint into the **Authorization Endpoint** field, and the token endpoint into the **Token Endpoint** and **Refresh Token Endpoint** fields.
  5. Add the following parameters to your resource:
  * Resource parameters
   + grant_type: **Token Value** = authorization_code, **Refresh Value** = refresh_token
   + client_id: the app ID that you made a note of for all three values
   + redirect_uri: for **Authorization Value** and **Token Value** put in the URL https://{KUID}.onk2.com/Identity/token/oauth/2
   + scope: for **Authorization Value** and **Token Value** put TeamsApp.ReadWrite.All
   + response_type: for **Authorization Value** put code
   + client_secret: for **Token Value** and **Refresh Value** put your client secret
  6. Save the resource and confirm that it's similar to the following figure

  ![Example OAuth Resource for Microsoft Teams](/OAuthResource.png)

  Once you've created your OAuth resource, you can now use it to configure a service instance, but you must first create your service type.

  ### Creating your Microsoft Teams Connector Service Type

  Either build the project by cloning the repo or download the index.jssp from the release.

  Once you have your index.jssp file for the Microsoft Teams Connector, create your service type using the example form from the Help topic or one you've built yourself.

  When your service type is created, register a service instance using the OAuth resource you created, and generate SmartObjects.

## Creating Custom Workflow Steps

Also included in this repo are four SVG icons that you can use if you want to create custom workflow steps in the K2 Workflow Designer. The following icons are available:
+ Create Team (CreateTeam32.svg)
+ Update Team (UpdateTeam32.svg)
+ Create Channel (CreateChannel32.svg)
+ Send Message to Channel (MessageChannel32.svg)

For more information about creating custom workflow steps, see [Creating custom Workflow steps](https://help.k2.com/onlinehelp/k2cloud/DevRef/current/default.htm#Extend/WF/Steps/Steps-Creating.htm%3FTocPath%3DExtending%2520K2%2520Cloud%7CCustom%2520workflow%2520steps%7C_____3)

## Building the Microsoft Teams Connector

If you've made changes to the source code, use these steps to build a new version of your connector.

**Note**: This template requires [Node.js](https://nodejs.org/) v12.14.1+ to run.

Install the dependencies and devDependencies:

```bash
npm install
```

See the documentation for [@k2oss/k2-broker-core](https://www.npmjs.com/package/@k2oss/k2-broker-core)
for more information about how to use the broker SDK package.

# Running Unit Tests
To run the unit tests, run:

```bash
npm test
```

You can also use a development build for debugging and coverage gutters:

```bash
npm run test:dev
```

You will find the code coverage results in [coverage/index.html](./coverage/index.html).

# Building your bundled JS
When you're ready to build your connector, run the following command:

```bash
npm run build
```

Add :dev to the end of the build if you want to be able to read your bundled .js:

```bash
npm run build:dev
```

Find the bundled JavaScript file in the [dist/index.js](./dist/index.js). If you're using the example Create or Update From File SmartForm or you've built your own, rename this file to index.jssp in order to attach it to the form.

# Creating a service type
Once you have a bundled .js file, register the service type using the system SmartObject located
at System > Management > SmartObjects > SmartObjects > JavaScript Service
Provider and run the Create From File method. Note that you'll need a form to do this. You can build
your own or use the one provided in the Help topic.

### License

MIT, found in the [LICENSE](./LICENSE) file.

[www.k2.com](https://www.k2.com)