---
page_type: sample
description: This sample demonstrates how to use the Microsoft Graph JavaScript SDK to access data in Office 365 from SharePoint Framework (SPFX) applications.
products:
- ms-graph
- microsoft-graph-calendar-api
- office-exchange-online
languages:
- typescript
---

# Microsoft Graph sample SharePoint Framework app

This sample demonstrates how to use the Microsoft Graph JavaScript SDK to access data in Office 365 from SharePoint Framework (SPFX) applications.

## Prerequisites

Before you start this tutorial, you should have the following tools installed on your development machine.

- [Node.js](https://nodejs.org/en/download/releases/)
- [Yeoman](https://yeoman.io/)
- [Gulp](https://gulpjs.com/)
- [Yeoman SharePoint generator](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/yeoman-generator-for-spfx-intro)

You can find more details about requirements for SharePoint Framework development at [Set up your SharePoint Framework development environment](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment).

You should also have a Microsoft work or school account, with access to a global administrator account in the same organization. If you don't have a Microsoft account, you can [sign up for the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) to get a free Microsoft 365 subscription.

Your Microsoft 365 tenant should be [setup for SharePoint Framework development](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant), with an app catalog and testing site created before you start this tutorial.

> [!NOTE]
> This tutorial was written with the following versions of the above tools. The steps in this guide may work with other versions, but that has not been tested.
>
> - Node.js 8.11.0
> - Yeoman 4.3.0
> - Gulp 2.3.0
> - Yeoman SharePoint generator 1.15.2

## Deploy the web part

1. Run the following two commands in your CLI to build and package your web part.

    ```Shell
    gulp bundle --ship
    gulp package-solution --ship
    ```

1. Open your browser and go to your tenant's [SharePoint App Catalog](https://docs.microsoft.com/sharepoint/use-app-catalog). Select the **Apps for SharePoint** menu item on the left-hand side.

1. Upload the **./sharepoint/solution/graph-tutorial.sppkg** file.

1. In the **Do you trust...** prompt, confirm that the prompt lists the 4 Microsoft Graph permissions you set in the **package-solution.json** file. Select **Make this solution available to all sites in the organization**, then select **Deploy**.

1. Go to the [SharePoint admin center](https://admin.microsoft.com/sharepoint?page=classicfeatures&modern=true) using a tenant administrator.

1. In the left-hand menu, select **Advanced**, then **API access**.

1. Select each of the pending requests from the **graph-tutorial-client-side-solution** package and choose **Approve**.

### Test the web part

1. Go to a SharePoint site where you want to test the web part. Create a new page to test the web part on.

1. Use the web part picker to find the **GraphTutorial** web part and add it to the page.

1. The access token is printed below the **Welcome to SharePoint!** message in the web part. You can copy this token and parse it at [https://jwt.ms/](https://jwt.ms/) to confirm that it contains the permission scopes required by the web part.

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
