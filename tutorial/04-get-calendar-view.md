<!-- markdownlint-disable MD002 MD041 -->

The SharePoint Framework provides the [MSGraphClient](https://docs.microsoft.com/javascript/api/sp-http/msgraphclient?view=sp-typescript-latest) for making calls to Microsoft Graph. This class wraps the [Microsoft Graph JavaScript Client Library](https://github.com/microsoftgraph/msgraph-sdk-javascript), pre-authenticating it with the current logged on user.

Because it wraps the existing JavaScript library, its usage is the same, and it's fully compatible with the Microsoft Graph TypeScript definitions.

## Get the user's calendar

1. Open **./src/webparts/graphTutorial/GraphTutorialWebPart.ts** and add the following `import` statements at the top of the file.

    ```typescript
    import { MSGraphClient } from '@microsoft/sp-http';
    import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
    import { startOfWeek, endOfWeek, setDay } from 'date-fns';
    ```

1. Add the following function to the **GraphTutorialWebPart** class to render an error.

    :::code language="typescript" source="../demo/graph-tutorial/src/webparts/graphTutorial/GraphTutorialWebPart.ts" id="renderErrorSnippet":::

1. Add the following function to print out the events in the user's calendar.

    ```typescript
    private renderCalendarView(events: MicrosoftGraph.Event[]) : void {
      const viewContainer = this.domElement.querySelector('#calendarView');
      let html = '';

      // Temporary: print events as a list
      for(const event of events) {
        html += `
          <p class="${ styles.description }">Subject: ${event.subject}</p>
          <p class="${ styles.description }">Organizer: ${event.organizer.emailAddress.name}</p>
          <p class="${ styles.description }">Start: ${event.start.dateTime}</p>
          <p class="${ styles.description }">End: ${event.end.dateTime}</p>
          `;
      }

      viewContainer.innerHTML = html;
    }
    ```

1. Replace the existing `render` function with the following.

    :::code language="typescript" source="../demo/graph-tutorial/src/webparts/graphTutorial/GraphTutorialWebPart.ts" id="renderSnippet":::

    Notice what this code does.

    - It uses `this.context.msGraphClientFactory.getClient` to get an authenticated **MSGraphClient** object.
    - It calls the `/me/calendarView` endpoint, setting the `startDateTime` and `endDateTime` query parameters to the start and end of the current week.
    - It uses `select` to limit which fields are returned, requesting only the fields the app uses.
    - It uses `orderby` to sort the events by their start time.
    - It uses `top` to limit the results to 25 events.

## Deploy the web part

1. Run the following two commands in your CLI to build and package your web part.

    ```Shell
    gulp bundle --ship
    gulp package-solution --ship
    ```

1. Open your browser and go to your tenant's SharePoint App Catalog. Select the **Apps for SharePoint** menu item on the left-hand side.

1. Upload the **./sharepoint/solution/graph-tutorial.sppkg** file.

1. In the **Do you trust...** prompt, confirm that the prompt lists the 4 Microsoft Graph permissions you set in the **package-solution.json** file. Select **Make this solution available to all sites in the organization**, then select **Deploy**.

1. If you have not already approved the Graph permissions for your web part, do that now.

    1. Go to the [SharePoint admin center](https://admin.microsoft.com/sharepoint?page=classicfeatures&modern=true) using a tenant administrator.

    1. In the left-hand menu, select **Advanced**, then **API access**.

    1. Select each of the pending requests from the **graph-tutorial-client-side-solution** package and choose **Approve**.

        ![A screenshot of the SharePoint admin center's API access page](images/api-access.png)

## Test the web part

1. Go to a SharePoint site where you want to test the web part. Create a new page to test the web part on.

1. Use the web part picker to find the **GraphTutorial** web part and add it to the page.

    ![A screenshot of the GraphTutorial web part in the web part picker](images/add-web-part.png)

1. A list of events for the current week are printed in the web part.

    ![A screenshot of the web part displaying a list of events](images/calendar-list.png)
