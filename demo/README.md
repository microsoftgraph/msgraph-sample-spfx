# How to run the completed project

## Prerequisites

To run the completed project in this folder, you need the following:

- [Node.js](https://nodejs.org/en/download/releases/) version 10.x
- [Gulp](https://gulpjs.com/)
- A Microsoft work or school account, with access to a global administrator account in the same organization. If you don't have a Microsoft account, you can [sign up for the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) to get a free Microsoft 365 subscription.
- Your Microsoft 365 tenant should be [setup for SharePoint Framework development](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant), with an app catalog and testing site created before you start this tutorial.

### Deploy the web part

1. Run the following two commands in your CLI to build and package your web part.

    ```Shell
    gulp bundle --ship
    gulp package-solution --ship
    ```

1. Open your browser and go to your tenant's SharePoint App Catalog. Select the **Apps for SharePoint** menu item on the left-hand side.

1. Upload the **./sharepoint/solution/graph-tutorial.sppkg** file.

1. In the **Do you trust...** prompt, confirm that the prompt lists the 4 Microsoft Graph permissions you set in the **package-solution.json** file. Select **Make this solution available to all sites in the organization**, then select **Deploy**.

1. Go to the [SharePoint admin center](https://admin.microsoft.com/sharepoint?page=classicfeatures&modern=true) using a tenant administrator.

1. In the left-hand menu, select **Advanced**, then **API access**.

1. Select each of the pending requests from the **graph-tutorial-client-side-solution** package and choose **Approve**.

    ![A screenshot of the SharePoint admin center's API access page](../tutorial/images/api-access.png)

### Test the web part

1. Go to a SharePoint site where you want to test the web part. Create a new page to test the web part on.

1. Use the web part picker to find the **GraphTutorial** web part and add it to the page.

    ![A screenshot of the GraphTutorial web part in the web part picker](../tutorial/images/add-web-part.png)
