# Leveraging the Microsoft Graph JavaScript SDK from the SharePoint Framework

This module will introduce you to working with the Microsoft Graph REST API to access data in Office 365 using the SharePoint Framework (SPFx). You will build three SPFx client-side web parts in a single SharePoint framework solution.

## In this lab

* [Show profile details from Microsoft Graph in SPFx client-side web part](#exercise1)
* [Show calendar events from Microsoft Graph in SPFx client-side web part](#exercise2)
* [Show Planner tasks from Microsoft Graph in SPFx client-side web part](#exercise3)

## Prerequisites

To complete this lab, you need the following:

* Office 365 tenancy
  * If you do not have one, you obtain one (for free) by signing up to the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).
* SharePoint Online environment
  * Refer to the SharePoint Framework documentation, specifically the **[Getting Started > Set up Office 365 Tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)** for the most current steps
* Local development environment with the latest 
  * Refer to the SharePoint Framework documentation, specifically the **[Getting Started > Set up development environment](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)** for the most current steps

<a name="exercise1"></a>

## Exercise 1: Show profile details from Microsoft Graph in SPFx client-side web part

In this exercise you will create a new SPFx project with a single client-side web part that uses React, [Fabric React](https://developer.microsoft.com/fabric) and the Microsoft Graph to display the currently logged in user's personal details in a familiar office [Persona](https://developer.microsoft.com/fabric#/components/persona) card.

### Create the SPFx Solution

1. Open a command prompt and change to the folder where you want to create the project.
1. Run the SharePoint Yeoman generator by executing the following command:

    ```shell
    yo @microsoft/sharepoint --plusbeta
    ```

    Use the following to complete the prompt that is displayed:

    * **What is your solution name?**: MSGraphSPFx
    * **Which baseline packages do you want to target for your component(s)?**: SharePoint Online only (latest)
    * **Where do you want to place the files?**: Use the current folder
    * **Do you want to allow the tenant admin the choice of being able to deploy the solution to all sites immediately without running any feature deployment or adding apps in sites?**: No
    * **Which type of client-side component to create?**: WebPart
    * **What is your Web part name?**: GraphPersona
    * **What is your Web part description?**: Display current user's persona details in a Fabric React Persona card
    * **Which framework would you like to use?**: React

    After provisioning the folders required for the project, the generator will install all the dependency packages using NPM.

1. When NPM completes downloading all dependencies, open the project in Visual Studio Code.

### Update Solution Dependencies

1. Install the Microsoft Graph Typescript type declarations by executing the following statement on the command line:

    ```shell
    npm install @microsoft/microsoft-graph-types --save-dev
    ```

1. The web part will use the Fabric React controls to display user interface components. Configure the project to use Fabric React:
    1. Execute the following on the command line to uninstall the SPFx Fabric Core library which is not needed as it is included in Fabric React:

        ```shell
        npm uninstall @microsoft/sp-office-ui-fabric-core
        ```

    1. Configure the included components styles to use the Fabric Core CSS from the Fabric React project.
        1. Open the **src\webparts\graphPersona\components\GraphPersona.module.scss**
        1. Replace the first line:

            ```css
            @import '~@microsoft/sp-office-ui-fabric-core/dist/sass/SPFabricCore.scss';
            ```

            With the following:

            ```css
            @import '~office-ui-fabric-react/dist/sass/_References.scss';
            ```

### Update the Web Part

Update the default web part to pass into the React component an instance of the Microsoft Graph client API:

1. Open the web part file **src\webparts\graphPersona\GraphPersonaWebPart.ts**.
1. Add the following `import` statements after the existing `import` statements:

    ```ts
    import { MSGraphClient } from '@microsoft/sp-client-preview';
    import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
    ```

1. Locate the `render()` method. This method creates a new instance of a React element by passing in the component class and the properties to bind to it. The only property being set is the `description` property.

    Replace the contents of the `render()` method with the following code to create an initialize a new instance fo the Microsoft Graph client:

    ```ts
    const element: React.ReactElement<IGraphPersonaProps> = React.createElement(
      GraphPersona,
      {
        graphClient: this.context.serviceScope.consume(MSGraphClient.serviceKey)
      }
    );

    ReactDom.render(element, this.domElement);
    ```

### Implement the React Component

1. After updating the public signature of the **GraphPersona** component, the public property interface of the component needs to be updated to accept the Microsoft Graph client:
    1. Open the **src\webparts\graphPersona\components\IGraphPersonaProps.tsx**
    1. Replace the contents with the following code to change the public signature of the component:

        ```ts
        import { MSGraphClient } from '@microsoft/sp-client-preview';

        export interface IHelloWorldProps {
          graphClient: MSGraphClient;
        }
        ```

1. Create a new interface that will keep track of the state of the component's state:
    1. Create a new file **IGraphPersonaState.ts** and save it to the folder: **src\webparts\graphResponse\components\**.
    1. Add the following code to define a new state object that will be used by the component:

        ```ts
        import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

        export interface IGraphPersonaProps {
          name: string;
          email: string;
          phone: string;
          image: string;
        }
        ```

1. Update the component's references to add the new state interface, support for the Microsoft Graph, Fabric React Persona control and other necessary controls.
    1. Open the **src\webparts\graphPersona\components\GraphPersona.tsx**
    1. Add the following `import` statements after the existing `import` statements:

        ```ts
        import { IGraphPersonaState } from './IGraphPersonaState';

        import { MSGraphClient } from '@microsoft/sp-client-preview';
        import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

        import {
          Persona,
          PersonaSize
        } from 'office-ui-fabric-react/lib/components/Persona';
        import { Link } from 'office-ui-fabric-react/lib/components/Link';
        ```

1. Update the public signature of the component to include the state:
    1. Locate the class `GraphPersona` declaration.
    1. At the end of the line, notice there is generic type with two parameters, the second is an empty object `{}`:

        ```ts
        export default class GraphPersona extends React.Component<IGraphPersonaProps, {}>
        ```

    1. Update the second parameter to be the state interface previously created:

        ```ts
        export default class GraphPersona extends React.Component<IGraphPersonaProps, IGraphPersonaState>
        ```

1. Add the following constructor to the `GraphPersona` class to initialize the state of the component:

    ```ts
    constructor(props: IGraphPersonaProps) {
      super(props);

      this.state = {
        name: '',
        email: '',
        phone: '',
        image: null
      };
    }
    ```

1. Add the Fabric React Persona card to the `render()` method's return statement:

    ```ts
    public render(): React.ReactElement<IGraphPersonaProps> {
      return (
        <Persona primaryText={this.state.name}
                secondaryText={this.state.email}
                onRenderSecondaryText={this._renderMail}
                tertiaryText={this.state.phone}
                onRenderTertiaryText={this._renderPhone}
                imageUrl={this.state.image}
                size={PersonaSize.size100} />
      );
    }
    ```

1. The code in the Persona card references two utility methods to control rendering of the secondary & tertiary text. Add the following to methods to the `GraphPersona` class that will be used to render the text accordingly:

    ```ts
    private _renderMail = () => {
      if (this.state.email) {
        return <Link href={`mailto:${this.state.email}`}>{this.state.email}</Link>;
      } else {
        return <div />;
      }
    }

    private _renderPhone = () => {
      if (this.state.phone) {
        return <Link href={`tel:${this.state.phone}`}>{this.state.phone}</Link>;
      } else {
        return <div />;
      }
    }
    ```

1. The last step is to update the loading, or *mounting* phase of the React component. When the component loads on the page, it should call the Microsoft Graph to get details on the current user as well as their photo. When each of these results complete, they will update the component's state which will trigger the component to rerender.

    Add the following method to the `GraphPersona` class:

    ```ts
    public componentDidMount(): void {
      this.props.graphClient
        .api(`me`)
        .get((error: any, user: MicrosoftGraph.User, rawResponse?: any) => {
          this.setState({
            name: user.displayName,
            email: user.mail,
            phone: user.businessPhones[0]
          });
        });

      this.props.graphClient
        .api('/me/photo/$value')
        .responseType('blob')
        .get((err: any, photoResponse: any, rawResponse: any) => {
          const blobUrl = window.URL.createObjectURL(rawResponse.xhr.response);
          this.setState({ image: blobUrl });
        });
    }
    ```

### Update the SPFx Package Permission Requests

The last step before testing is to notify SharePoint that upon deployment to production, this app requires permission to the Microsoft Graph to access the user's persona details.

1. Open the **config\package-solution.json** file.
1. Locate the `solution` section. Add the following permission request element just after the property `includeClientSideAssets`:

    ```json
    "webApiPermissionRequests": [
      {
        "resource": "Microsoft Graph",
        "scope": "User.ReadBasic.All"
      }
    ]
    ```

### Test the Solution

1. Create the SharePoint package for deployment:
    1. Build the solution by executing the following on the command line:

        ```shell
        gulp build
        ```

    1. Bundle the solution by executing the following on the command line:

        ```shell
        gulp bundle --ship
        ```

    1. Package the solution by executing the following on the command line:

        ```shell
        gulp package-solution --ship
        ```

1. Deploy and trust the SharePoint package:
    1. In the browser, navigate to your SharePoint Online Tenant App Catalog.

        >Note: Creation of the Tenant App Catalog site is one of the steps in the **[Getting Started > Set up Office 365 Tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)** setup documentation.

    1. Select the **Apps for SharePoint** link in the navigation:

        ![Screenshot of the navigation in the SharePoint Online Tenant App Catalog](./Images/tenant-app-catalog-01.png)

    1. Drag the generated SharePoint package from **\sharepoint\solution\ms-graph-sp-fx.sppkg** into the **Apps for SharePoint** library.
    1. In the **Do you trust ms-graph-sp-fx-client-side-solution?** dialog, select **Deploy**.

        ![Screenshot of trusting a SharePoint package](./Images/tenant-app-catalog-02.png)

1. Approve the API permission request:
    1. Navigate to the SharePoint Admin Portal located at **https://{{REPLACE_WITH_YOUR_TENANTID}}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx**, replacing the domain with your SharePoint Online's administration tenant URL.

        >Note: At the time of writing, this feature is only in the SharePoint Online preview portal.

    1. In the navigation, select **Advanced > API Management**:

        ![Screenshot of the SharePoint Online admin portal](./Images/spo-admin-portal-01.png)

    1. Select the **Pending approval** for the **Microsoft Graph** permission **User.ReadBasic.All**.

        ![Screenshot of the SharePoint Online admin portal API Management page](./Images/spo-admin-portal-02.png)

    1. Select the **Approve or Reject** button, followed by selecting **Approve**.

        ![Screenshot of the SharePoint Online permission approval](./Images/spo-admin-portal-03.png)

1. Test the web part:

    >NOTE: The SharePoint Framework includes a locally hosted & SharePoint Online hosted workbench for testing custom solutions. However, the workbench will not work the first time when testing solutions that utilize the Microsoft due to nuances with how the workbench operates and authentication requirements. Therefore, the first time you test a Microsoft Graph enabled SPFx solution, you will need to test them in a real modern page.
    >
    >Once this has been done and your browser has been cookied by the Azure AD authentication process, you can leverage local webserver and SharePoint Online-hosted workbench for testing the solution.

    1. Setup environment to test the web part on a real SharePoint Online modern page:

        1. In a browser, navigate to a SharePoint Online site.
        1. In the site navigation, select the **Pages** library.
        1. Select an existing page (*option 2 in the following image*), or create a new page (*option 1 in the following image*) in the library to test the web part on.

            ![Screenshot of the SharePoint Online Pages library](./Images/graph-test-01.png)

            *Continue with the test by skipping the next section to add the web part to the page.*

    1. Setup environment to test the from the local webserver and hosted workbench:
        1. In the command prompt for the project, execute the following command to start the local web server:

            ```shell
            gulp serve --nobrowser
            ```

        1. In a browser, navigate to one of your SharePoint Online site's hosted workbench located at **/_layouts/15/workbench.aspx**

    1. Add the web part to the page and test:
        1. In the browser, select the Web part icon button to open the list of available web parts:

            ![Screenshot of adding the web part to the hosted workbench](./Images/graph-persona-01.png)

        1. Locate the **GraphPersona** web part and select it

            ![Screenshot of adding the web part to the hosted workbench](./Images/graph-persona-02.png)

        1. When the page loads, notice after a brief delay, it will display the current user's details on the Persona card:

            ![Screenshot of the web part running in the hosted workbench](./Images/graph-persona-03.png)

<a name="exercise2"></a>

## Exercise 2: Show calendar events from Microsoft Graph in SPFx client-side web part

In this exercise you add a client-side web part that uses React, [Fabric React](https://developer.microsoft.com/fabric) and the Microsoft Graph to an existing SPFx project that will display a list of the current user's calendar events using the [List](https://developer.microsoft.com/fabric#/components/list) component.

> This exercise assumes you completed exercise 1 and created a SPFx solution and configured the project. If not, complete the section [Create the SPFx Solution](#create-the-spfx-solution).

## Add SPFx Component to Existing SPFx Solution

1. Open a command prompt and change to the folder of the existing SPFx solution.
1. Run the SharePoint Yeoman generator by executing the following command:

    ```shell
    yo @microsoft/sharepoint --plusbeta
    ```

    Use the following to complete the prompt that is displayed:

    * **Which type of client-side component to create?**: WebPart
    * **What is your Web part name?**: GraphEventsList
    * **What is your Web part description?**: Display current user's calendar events in a Fabric React List
    * **Which framework would you like to use?**: React

    After provisioning the folders required for the project, the generator will install all the dependency packages using NPM.

1. The project will use a library to assist in working with dates. Add this by executing the following command in the command prompt:

    ```shell
    npm install date-fns --save
    ```

1. Open the project in Visual Studio Code.

### Update the Web Part

Update the default web part to pass into the React component an instance of the Microsoft Graph client API:

1. Open the web part file **src\webparts\graphEventsList\GraphEventsListWebPart.ts**.
1. Add the following `import` statements after the existing `import` statements:

    ```ts
    import { MSGraphClient } from '@microsoft/sp-client-preview';
    import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
    ```

1. Locate the `render()` method. This method creates a new instance of a React element by passing in the component class and the properties to bind to it. The only property being set is the `description` property.

    Replace the contents of the `render()` method with the following code to create an initialize a new instance fo the Microsoft Graph client:

    ```ts
    const element: React.ReactElement<IGraphEventsListProps> = React.createElement(
      GraphEventsList,
      {
        graphClient: this.context.serviceScope.consume(MSGraphClient.serviceKey)
      }
    );

    ReactDom.render(element, this.domElement);
    ```

### Implement the React Component

1. After updating the public signature of the **GraphEventsList** component, the public property interface of the component needs to be updated to accept the Microsoft Graph client:
    1. Open the **src\webparts\graphEventsList\components\IGraphEventsListProps.tsx**
    1. Replace the contents with the following code to change the public signature of the component:

        ```ts
        import { MSGraphClient } from '@microsoft/sp-client-preview';

        export interface IGraphEventsListProps {
          graphClient: MSGraphClient;
        }
        ```

1. Create a new interface that will keep track of the state of the component's state:
    1. Create a new file **IGraphEventsListState.ts** and save it to the folder: **src\webparts\graphEventsList\components\**.
    1. Add the following code to define a new state object that will be used by the component:

        ```ts
        import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

        export interface IGraphEventsListState {
          events: MicrosoftGraph.Event[];
        }
        ```

1. Update the component's references to add the new state interface, support for the Microsoft Graph, Fabric React List and other necessary controls.
    1. Open the **src\webparts\graphEventsList\components\GraphEventsList.tsx**
    1. Add the following `import` statements after the existing `import` statements:

        ```ts
        import { IGraphEventsListState } from './IGraphEventsListState';

        import { MSGraphClient } from '@microsoft/sp-client-preview';
        import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

        import { List } from 'office-ui-fabric-react/lib/List';
        import { format } from 'date-fns';
        ```

1. Update the public signature of the component to include the state:
    1. Locate the class `GraphEventsList` declaration.
    1. At the end of the line, notice there is generic type with two parameters, the second is an empty object `{}`:

        ```ts
        export default class GraphEventsList extends React.Component<IGraphEventsListProps, {}>
        ```

    1. Update the second parameter to be the state interface previously created:

        ```ts
        export default class GraphEventsList extends React.Component<IGraphEventsListProps, IGraphEventsListState>
        ```

1. Add the following constructor to the `GraphEventsList` class to initialize the state of the component:

    ```ts
    constructor(props: IGraphEventsListProps) {
      super(props);

      this.state = {
        events: []
      };
    }
    ```

1. Add the Fabric React List to the `render()` method's return statement:

    ```ts
    public render(): React.ReactElement<IGraphEventsListProps> {
      return (
        <List items={this.state.events} 
              onRenderCell={this._onRenderEventCell} />
      );
    }
    ```

1. The code in the List card references a utility methods to control rendering of the list cell. Add the following to method to the `GraphEventsList` class that will be used to render the cell accordingly:

    ```ts
    private _onRenderEventCell(item: MicrosoftGraph.Event, index: number | undefined): JSX.Element {
      return (
        <div>
          <h3>{item.subject}</h3>
          {format( new Date(item.start.dateTime), 'MMMM Mo, YYYY h:mm A')} - {format( new Date(item.end.dateTime), 'h:mm A')}
        </div>
      );
    }
    ```

1. The last step is to update the loading, or *mounting* phase of the React component. When the component loads on the page, it should call the Microsoft Graph to get current user's calendar events. When each of these results complete, they will update the component's state which will trigger the component to rerender.

    Add the following method to the `GraphEventsList` class:

    ```ts
    public componentDidMount(): void {
      this.props.graphClient
        .api('/me/events')
        .get((error: any, eventsResponse: any, rawResponse?: any) => {
          const calendarEvents: MicrosoftGraph.Event[] = eventsResponse.value;
          console.log('calendarEvents', calendarEvents);
          this.setState({ events: calendarEvents });
        });
    }
    ```

### Update the SPFx Package Permission Requests

The last step before testing is to notify SharePoint that upon deployment to production, this app requires permission to the Microsoft Graph to access the user's calendar events.

1. Open the **config\package-solution.json** file.
1. Locate the `webApiPermissionRequests` property. Add the following permission request element just after the existing permission:

    ```json
    {
      "resource": "Microsoft Graph",
      "scope": "Calendars.Read"
    }
    ```

### Test the Solution

1. Create the SharePoint package for deployment:
    1. Build the solution by executing the following on the command line:

        ```shell
        gulp build
        ```

    1. Bundle the solution by executing the following on the command line:

        ```shell
        gulp bundle --ship
        ```

    1. Package the solution by executing the following on the command line:

        ```shell
        gulp package-solution --ship
        ```

1. Deploy and trust the SharePoint package:
    1. In the browser, navigate to your SharePoint Online Tenant App Catalog.

        >Note: Creation of the Tenant App Catalog site is one of the steps in the **[Getting Started > Set up Office 365 Tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)** setup documentation.

    1. Select the **Apps for SharePoint** link in the navigation:

        ![Screenshot of the navigation in the SharePoint Online Tenant App Catalog](./Images/tenant-app-catalog-01.png)

    1. Drag the generated SharePoint package from **\sharepoint\solution\ms-graph-sp-fx.sppkg** into the **Apps for SharePoint** library.
        1. If you previously uploaded the same package, as in the case from exercise 1, if the **A file with the same name already exists** dialog, select the **Replace It** button.
    1. In the **Do you trust ms-graph-sp-fx-client-side-solution?** dialog, select **Deploy**.

        ![Screenshot of trusting a SharePoint package](./Images/tenant-app-catalog-02.png)

1. Approve the API permission request:
    1. Navigate to the SharePoint Admin Portal located at **https://{{REPLACE_WITH_YOUR_TENANTID}}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx**, replacing the domain with your SharePoint Online's administration tenant URL.

        >Note: At the time of writing, this feature is only in the SharePoint Online preview portal.

    1. In the navigation, select **Advanced > API Management**:

        ![Screenshot of the SharePoint Online admin portal](./Images/spo-admin-portal-01.png)

    1. Select the **Pending approval** for the **Microsoft Graph** permission **Calendars.Read**.

        ![Screenshot of the SharePoint Online admin portal API Management page](./Images/spo-admin-portal-04.png)

    1. Select the **Approve or Reject** button, followed by selecting **Approve**.

        ![Screenshot of the SharePoint Online permission approval](./Images/spo-admin-portal-03.png)

1. Test the web part:

    >NOTE: The SharePoint Framework includes a locally hosted & SharePoint Online hosted workbench for testing custom solutions. However, the workbench will not work the first time when testing solutions that utilize the Microsoft due to nuances with how the workbench operates and authentication requirements. Therefore, the first time you test a Microsoft Graph enabled SPFx solution, you will need to test them in a real modern page.
    >
    >Once this has been done and your browser has been cookied by the Azure AD authentication process, you can leverage local webserver and SharePoint Online-hosted workbench for testing the solution.

    1. Setup environment to test the web part on a real SharePoint Online modern page:

        1. In a browser, navigate to a SharePoint Online site.
        1. In the site navigation, select the **Pages** library.
        1. Select an existing page (*option 2 in the following image*), or create a new page (*option 1 in the following image*) in the library to test the web part on.

            ![Screenshot of the SharePoint Online Pages library](./Images/graph-test-01.png)

            *Continue with the test by skipping the next section to add the web part to the page.*

    1. Setup environment to test the from the local webserver and hosted workbench:
        1. In the command prompt for the project, execute the following command to start the local web server:

            ```shell
            gulp serve --nobrowser
            ```

        1. In a browser, navigate to one of your SharePoint Online site's hosted workbench located at **/_layouts/15/workbench.aspx**

    1. Add the web part to the page and test:
        1. In the browser, select the Web part icon button to open the list of available web parts:

            ![Screenshot of adding the web part to the hosted workbench](./Images/graph-persona-01.png)

        1. Locate the **GraphEventList** web part and select it

            ![Screenshot of adding the web part to the hosted workbench](./Images/graph-eventList-01.png)

        1. When the page loads, notice after a brief delay, it will display the current user's calendar events in the list

            ![Screenshot of the web part running in the hosted workbench](./Images/graph-eventList-02.png)

<a name="exercise3"></a>

## Exercise 3: Show Planner tasks from Microsoft Graph in SPFx client-side web part

In this exercise you add a client-side web part that uses React, [Fabric React](https://developer.microsoft.com/fabric) and the Microsoft Graph to an existing SPFx project that will display a list of the current user's tasks from Planner using the [List](https://developer.microsoft.com/fabric#/components/list) component.

> This exercise assumes you completed exercise 1 and created a SPFx solution and configured the project. If not, complete the section [Create the SPFx Solution](#create-the-spfx-solution).

## Add SPFx Component to Existing SPFx Solution

1. Open a command prompt and change to the folder of the existing SPFx solution.
1. Run the SharePoint Yeoman generator by executing the following command:

    ```shell
    yo @microsoft/sharepoint --plusbeta
    ```

    Use the following to complete the prompt that is displayed:

    * **Which type of client-side component to create?**: WebPart
    * **What is your Web part name?**: GraphTasks
    * **What is your Web part description?**: Display current user's tasks from Planner in a Fabric React List
    * **Which framework would you like to use?**: React

    After provisioning the folders required for the project, the generator will install all the dependency packages using NPM.

1. Open the project in Visual Studio Code.

### Update the Web Part

Update the default web part to pass into the React component an instance of the Microsoft Graph client API:

1. Open the web part file **src\webparts\graphTasks\GraphTasksWebPart.ts**.
1. Add the following `import` statements after the existing `import` statements:

    ```ts
    import { MSGraphClient } from '@microsoft/sp-client-preview';
    import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
    ```

1. Locate the `render()` method. This method creates a new instance of a React element by passing in the component class and the properties to bind to it. The only property being set is the `description` property.

    Replace the contents of the `render()` method with the following code to create an initialize a new instance fo the Microsoft Graph client:

    ```ts
    const element: React.ReactElement<IGraphTasksProps> = React.createElement(
      GraphPersona,
      {
        graphClient: this.context.serviceScope.consume(MSGraphClient.serviceKey)
      }
    );

    ReactDom.render(element, this.domElement);
    ```

### Implement the React Component

1. After updating the public signature of the **GraphTasks** component, the public property interface of the component needs to be updated to accept the Microsoft Graph client:
    1. Open the **src\webparts\graphTasks\components\IGraphTasksProps.tsx**
    1. Replace the contents with the following code to change the public signature of the component:

        ```ts
        import { MSGraphClient } from '@microsoft/sp-client-preview';

        export interface IGraphTasksProps {
          graphClient: MSGraphClient;
        }
        ```

1. Create a new interface that will keep track of the state of the component's state:
    1. Create a new file **IGraphTasksState.ts** and save it to the folder: **src\webparts\graphTasks\components\**.
    1. Add the following code to define a new state object that will be used by the component:

        ```ts
        import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

        export interface IGraphTasksState {
          tasks: MicrosoftGraph.PlannerTask[];
        }
        ```

1. Update the component's references to add the new state interface, support for the Microsoft Graph, Fabric React List and other necessary controls.
    1. Open the **src\webparts\graphTasks\components\GraphTasks.tsx**
    1. Add the following `import` statements after the existing `import` statements:

        ```ts
        import { IGraphTasksState } from './IGraphTasksState';

        import { MSGraphClient } from '@microsoft/sp-client-preview';
        import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

        import { List } from 'office-ui-fabric-react/lib/List';
        import { format } from 'date-fns';
        ```

1. Update the public signature of the component to include the state:
    1. Locate the class `GraphTasks` declaration.
    1. At the end of the line, notice there is generic type with two parameters, the second is an empty object `{}`:

        ```ts
        export default class GraphTasks extends React.Component<IGraphTasksProps, {}>
        ```

    1. Update the second parameter to be the state interface previously created:

        ```ts
        export default class GraphTasks extends React.Component<IGraphTasksProps, IGraphTasksState>
        ```

1. Add the following constructor to the `GraphTasks` class to initialize the state of the component:

    ```ts
    constructor(props: IGraphTasksProps) {
      super(props);

      this.state = {
        tasks: []
      };
    }
    ```

1. Add the Fabric React List to the `render()` method's return statement:

    ```ts
    public render(): React.ReactElement<IGraphTasksProps> {
      return (
        <List items={this.state.tasks}
              onRenderCell={this._onRenderEventCell} />
      );
    }
    ```

1. The code in the List card references a utility methods to control rendering of the list cell. Add the following to method to the `GraphTasks` class that will be used to render the cell accordingly:

    ```ts
    private _onRenderEventCell(item: MicrosoftGraph.PlannerTask, index: number | undefined): JSX.Element {
      return (
        <div>
          <h3>{item.subject}</h3>
          <strong>Due:</strong> {format( new Date(item.dueDateTime), 'MMMM Mo, YYYY at h:mm A')}
        </div>
      );
    }
    ```

1. The last step is to update the loading, or *mounting* phase of the React component. When the component loads on the page, it should call the Microsoft Graph to get current user's calendar events. When each of these results complete, they will update the component's state which will trigger the component to rerender.

    Add the following method to the `GraphEventsList` class:

    ```ts
    public componentDidMount(): void {
      this.props.graphClient
        .api('/me/planner/tasks')
        .get((error: any, tasksResponse: any, rawResponse?: any) => {
          console.log('tasksResponse', tasksResponse);
          const plannerTasks: MicrosoftGraph.PlannerTask[] = tasksResponse.value;
          this.setState({ tasks: plannerTasks });
        });
    }
    ```

### Update the SPFx Package Permission Requests

The last step before testing is to notify SharePoint that upon deployment to production, this app requires permission to the Microsoft Graph to access the user's calendar events.

1. Open the **config\package-solution.json** file.
1. Locate the `webApiPermissionRequests` property. Add the following permission request element just after the existing permission:

    ```json
    {
      "resource": "Microsoft Graph",
      "scope": "Group.Read.All"
    }
    ```

    >Note: There are multiple *"task"* related permissions (scopes) used with the Microsoft Graph. Planner tasks are accessible via the `Groups.Read.All` scope while Outlook/Exchange tasks are accessible via the `Tasks.Read` scope.

### Test the Solution

1. Create the SharePoint package for deployment:
    1. Build the solution by executing the following on the command line:

        ```shell
        gulp build
        ```

    1. Bundle the solution by executing the following on the command line:

        ```shell
        gulp bundle --ship
        ```

    1. Package the solution by executing the following on the command line:

        ```shell
        gulp package-solution --ship
        ```

1. Deploy and trust the SharePoint package:
    1. In the browser, navigate to your SharePoint Online Tenant App Catalog.

        >Note: Creation of the Tenant App Catalog site is one of the steps in the **[Getting Started > Set up Office 365 Tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)** setup documentation.

    1. Select the **Apps for SharePoint** link in the navigation:
    1. Drag the generated SharePoint package from **\sharepoint\solution\ms-graph-sp-fx.sppkg** into the **Apps for SharePoint** library.
        1. If you previously uploaded the same package, as in the case from exercise 1, if the **A file with the same name already exists** dialog, select the **Replace It** button.
    1. In the **Do you trust ms-graph-sp-fx-client-side-solution?** dialog, select **Deploy**.
1. Approve the API permission request:
    1. Navigate to the SharePoint Admin Portal located at **https://{{REPLACE_WITH_YOUR_TENANTID}}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx**, replacing the domain with your SharePoint Online's administration tenant URL.

        >Note: At the time of writing, this feature is only in the SharePoint Online preview portal.

    1. In the navigation, select **Advanced > API Management**:
    1. Select the **Pending approval** for the **Microsoft Graph** permission **Group.Read.All**.
    1. Select the **Approve or Reject** button, followed by selecting **Approve**.

1. Test the web part:

    >NOTE: The SharePoint Framework includes a locally hosted & SharePoint Online hosted workbench for testing custom solutions. However, the workbench will not work the first time when testing solutions that utilize the Microsoft due to nuances with how the workbench operates and authentication requirements. Therefore, the first time you test a Microsoft Graph enabled SPFx solution, you will need to test them in a real modern page.
    >
    >Once this has been done and your browser has been cookied by the Azure AD authentication process, you can leverage local webserver and SharePoint Online-hosted workbench for testing the solution.

    1. Setup environment to test the web part on a real SharePoint Online modern page:

        1. In a browser, navigate to a SharePoint Online site.
        1. In the site navigation, select the **Pages** library.
        1. Select an existing page (*option 2 in the following image*), or create a new page (*option 1 in the following image*) in the library to test the web part on.

            ![Screenshot of the SharePoint Online Pages library](./Images/graph-test-01.png)

            *Continue with the test by skipping the next section to add the web part to the page.*

    1. Setup environment to test the from the local webserver and hosted workbench:
        1. In the command prompt for the project, execute the following command to start the local web server:

            ```shell
            gulp serve --nobrowser
            ```

        1. In a browser, navigate to one of your SharePoint Online site's hosted workbench located at **/_layouts/15/workbench.aspx**

    1. Add the web part to the page and test:
        1. In the browser, select the Web part icon button to open the list of available web parts:

            ![Screenshot of adding the web part to the hosted workbench](./Images/graph-persona-01.png)

        1. Locate the **GraphTasks** web part and select it

            ![Screenshot of adding the web part to the hosted workbench](./Images/graph-taskList-01.png)

        1. When the page loads, notice after a brief delay, it will display the current user's tasks in the list:

            ![Screenshot of the web part running in the hosted workbench](./Images/graph-taskList-02.png)
