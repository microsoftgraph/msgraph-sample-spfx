<!-- markdownlint-disable MD002 MD041 -->

In this tutorial, you'll create a [SharePoint client-side web part](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/overview-client-side-web-parts) that will use Microsoft Graph to get the user's calendar for the current week and allow the user to add an event to their calendar.

## Create a web part project

1. Open your command-line interface (CLI) in an empty directory where you want to create the project. Run the following command to start the Yeoman SharePoint generator.

    ```Shell
    yo @microsoft/sharepoint
    ```

1. Respond to the prompts as follows.

    - **What is your solution name?** `graph-tutorial`
    - **Which baseline packages do you want to target for your component(s)?** `SharePoint Online only (latest)`
    - **Where do you want to place the files?** `Use the current folder`
    - **Do you want to allow the tenant admin the choice of being able to deploy the solution to all sites immediately without running any feature deployment or adding apps in sites?** `Yes`
    - **Will the components in the solution require permissions to access web APIs that are unique and not shared with other components in the ten
    ant?** `No`
    - **Which type of client-side component to create?** `WebPart`
    - **What is your Web part name?** `GraphTutorial`
    - **What is your Web part description?** `GraphTutorial description`
    - **Which framework would you like to use?** `No JavaScript framework`

1. Run the following command to update the TypeScript version in the project to 3.7.

    ```Shell
    npm install @microsoft/rush-stack-compiler-3.7 --save-dev
    ```

1. Open **./tsconfig.json** and replace `rush-stack-compiler-3.3` with `rush-stack-compiler-3.7`.

1. Open **./tslint.json** and remove the `"no-use-before-declare": true,` line. The `no-use-before-declare` rule is deprecated and will cause an error during the build process.

## Install dependencies

Before moving on, install some additional NPM packages that you will use later.

- [Microsoft Graph TypeScript definitions](https://github.com/microsoftgraph/msgraph-typescript-typings) to provide Intellisense for Microsoft Graph objects.
- [Microsoft Graph Toolkit](https://docs.microsoft.com/graph/toolkit/overview) to provide UI components for the web part.
- [date-fns](https://date-fns.org/) for helpful functions for working with dates.

```Shell
npm install @microsoft/microsoft-graph-types@1.16.0 --save-dev
npm install @microsoft/mgt@1.3.4 date-fns @2.15.0
```
