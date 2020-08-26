<!-- markdownlint-disable MD002 MD041 -->

In this section, you'll use the [Microsoft Graph Toolkit](https://docs.microsoft.com/graph/toolkit/overview) to replace the simple list of events with rich UI.

The toolkit provides an [Agenda component](https://docs.microsoft.com/graph/toolkit/components/agenda), which is well-suited to render our list of events.

## Update the web part

1. Open **./src/webparts/graphTutorial/GraphTutorialWebPart.module.scss**. Change the value of the `background-color` attribute in the `.row` entry to `$ms-color-white`.

    :::code language="css" source="../demo/graph-tutorial/src/webparts/graphTutorial/GraphTutorialWebPart.module.scss" id="rowScssSnippet" highlight="4":::

1. Add the following entry inside the `.graphTutorial` entry.

    :::code language="css" source="../demo/graph-tutorial/src/webparts/graphTutorial/GraphTutorialWebPart.module.scss" id="addSocialBtnSnippet":::

1. Open **./src/webparts/graphTutorial/GraphTutorialWebPart.ts** and add the following `import` statement at the top of the file.

    ```typescript
    import { Providers, SharePointProvider, MgtAgenda } from '@microsoft/mgt';
    ```

1. Add the following function to the **GraphTutorialWebPart** class to initialize the toolkit.

    :::code language="typescript" source="../demo/graph-tutorial/src/webparts/graphTutorial/GraphTutorialWebPart.ts" id="onInitSnippet":::

1. Replace the existing `renderCalendarView` function with the following.

    :::code language="typescript" source="../demo/graph-tutorial/src/webparts/graphTutorial/GraphTutorialWebPart.ts" id="renderCalendarViewSnippet":::

    This replaces the basic list with the **Agenda** component from the toolkit.

1. Build, package, and re-upload the web part, then refresh the page where you are testing it.

    ![A screenshot of the web part with the Agenda component](images/mgt-agenda.md)

## An alternate approach

At this point, you have code that:

- Uses the **MSGraphClient** to get the user's calendar view for the current week from Microsoft Graph.
- Add those events to the **Agenda** component from the Microsoft Graph Toolkit.

With this approach, you have full control over the Graph API call and can do any processing of the events prior to rendering that you want. However, if that isn't required, you can simplify by letting the **Agenda** component do the work for you.

All Microsoft Graph Toolkit components are capable of making all of the relevant API calls to the Microsoft Graph. For example, by just adding the **Agenda** component to the web part, and not setting any properties, the web part would use its default settings to get events for the current day. Let's look at how we can achieve the same results we currently have (events for the current week).

1. Replace the existing `render` method with the following.

    :::code language="typescript" source="../demo/graph-tutorial/src/webparts/graphTutorial/GraphTutorialWebPart.ts" id="alternateRenderSnippet":::

    Now, instead of making an API call in `render`, you simply add an `mgt-agenda` element directly into the HTML. By setting `date` to the start of the week, and `days` to 7, the component will make the same API call the previous version of `render` was making.

1. Add the following empty function to the **GraphTutorialWebPart** class.

    ```typescript
    private async addSocialToCalendar() {}
    ```

    > [!NOTE]
    > We also added an **Add team social** button to the web part, and added the `addSocialToCalendar` method as an event listener.  You'll implement the code behind that in the next section. For now, we just want the code to compile.

1. Build, package, and re-upload the web part, then refresh the page where you are testing it. The view should be the same as your previous test.

### Using the toolkit vs making API calls

At this point you may be wondering why you went through the trouble of using the **MSGraphClient** at all, when the toolkit does the work for you. The toolkit is designed for rendering results that you query from Microsoft Graph, such as a list of events. However, there are scenarios where making API calls yourself is necessary.

- Any API calls that are not a `GET` request. For example, creating a new event on the calendar, or updating a user's phone number.
- API calls to get data that's intended to be used "behind the scenes" and not rendered directly.
