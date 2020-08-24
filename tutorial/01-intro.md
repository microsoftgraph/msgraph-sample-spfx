<!-- markdownlint-disable MD002 MD041 -->

This tutorial teaches you how to build a [SharePoint client-side web part](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/overview-client-side-web-parts) that uses the Microsoft Graph API to retrieve calendar information for a user.

> [!TIP]
> If you prefer to just download the completed tutorial, you can download or clone the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-spfx). See the README file in the **demo** folder for instructions on configuring the app with an app ID.

## Prerequisites

Before you start this tutorial, you should have the following tools installed on your development machine.

- [Node.js](https://nodejs.org/en/download/releases/)
- [Yeoman](https://yeoman.io/)
- [Gulp](https://gulpjs.com/)
- [Yeoman SharePoint generator](https://docs.microsoft.com/sharepoint/dev/spfx/toolchain/scaffolding-projects-using-yeoman-sharepoint-generator)

You can find more details about requirements for SharePoint Framework development at [Set up your SharePoint Framework development environment](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment).

> [!IMPORTANT]
> SharePoint Framework requires Node.js version 10.x. Be sure to install the correct version.

You should also have a Microsoft work or school account, with access to a global administrator account in the same organization. If you don't have a Microsoft account, you can [sign up for the Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get a free Office 365 subscription.

Your Microsoft 365 tenant should be [setup for SharePoint Framework development](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant), with an app catalog and testing site created before you start this tutorial.

> [!NOTE]
> This tutorial was written with the following versions of the above tools. The steps in this guide may work with other versions, but that has not been tested.
>
> - Node.js
> - Yeoman
> - Gulp
> - Yeoman Sharepoint generator

## Feedback

Please provide any feedback on this tutorial in the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-spfx).
