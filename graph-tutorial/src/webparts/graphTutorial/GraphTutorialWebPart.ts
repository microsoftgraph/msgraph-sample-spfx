// Copyright (c) Microsoft Corporation.
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { startOfWeek, endOfWeek, setDay, set } from 'date-fns';
import { Providers, SharePointProvider, MgtAgenda } from '@microsoft/mgt';
import { MSGraphClientV3 } from '@microsoft/sp-http';


import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GraphTutorialWebPart.module.scss';
import * as strings from 'GraphTutorialWebPartStrings';

export interface IGraphTutorialWebPartProps {
  description: string;
}

export default class GraphTutorialWebPart extends BaseClientSideWebPart<IGraphTutorialWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  // <onInitSnippet>
  protected async onInit(): Promise<void>  {
    // Set the toolkit's global auth provider to
    // SharePoint provider, allowing it to use the Graph
    // access token
    Providers.globalProvider =
      new SharePointProvider(this.context);
  }
  // </onInitSnippet>

  // <renderSnippet>
  public render(): void {
    this.context.msGraphClientFactory
      .getClient('3')
      .then((graphClient: MSGraphClientV3)=> {
        // Get current date
        const now = new Date();
        // Get the start and end of the week based on current date
        const weekStart = startOfWeek(now);
        const weekEnd = endOfWeek(now);
        graphClient
          .api(`/me/calendarView?startDateTime=${weekStart.toISOString()}&endDateTime=${weekEnd.toISOString()}`)
          .select('subject,organizer,start,end,location,attendees')
          .orderby('start/dateTime')
          .top(25)
          .get((error: any, events: any) => { // eslint-disable-line
            this.domElement.innerHTML = `
            <div class="${ styles.graphTutorial }">
              <div class="${ styles.container }">
                <div class="${ styles.row }">
                  <div class="${ styles.column }">
                    <div id="calendarView" />
                  </div>
                </div>
              </div>
            </div>`;
            if (error) {
              this.renderGraphError(error);
            } else {
              this.renderCalendarView(events.value);
            }
          })
          .catch((reason: Error) => {
            this.renderGraphError(reason);
          });  
      })
      .catch((reason: Error) => {
        this.renderGraphError(reason);
      });  
  }

  // <renderGraphErrorSnippet>
  private renderGraphError(error: any): void { // eslint-disable-line
    const viewContainer = this.domElement.querySelector('#calendarView'); 

    // Basic error display
    viewContainer.innerHTML = `
    <h2 class="${ styles.title }">Error</h2>
    <code style="word-break: break-all;">${JSON.stringify(error, null, 2)}</code>`;
  }
  // </renderGraphErrorSnippet>

  // <renderCalendarViewSnippet>
  private renderCalendarView(events: MicrosoftGraph.Event[]) : void {
    const viewContainer = this.domElement.querySelector('#calendarView');

    // Create an agenda component from the toolkit
    const agenda = new MgtAgenda();
    // Set the events
    agenda.events = events;
    // Group events by day
    agenda.groupByDay = true;

    viewContainer.appendChild(agenda);
  }
  // </renderCalendarViewSnippet>

    // <addSocialToCalendarSnippet>
    private async addSocialToCalendar(): Promise<void>   {
      const graphClient = await this.context.msGraphClientFactory.getClient('3');
  
      // Get current date
      const now = new Date();
  
      // Set start time to next Friday
      // at 4 PM
      const socialHourStart = set(
        setDay(now, 5),
        {
          hours: 16,
          minutes: 0,
          seconds:0,
          milliseconds: 0
        });
  
      // Create a new event
      const socialHour: MicrosoftGraph.Event = {
        subject: 'Team Social Hour',
        body: {
          contentType: 'text',
          content: 'Come join the rest of the team for our end-of-week social hour.'
        },
        location: {
          displayName: 'Break room'
        },
        start: {
          dateTime: socialHourStart.toISOString(),
          timeZone: 'UTC'
        },
        end: {
          dateTime: set(socialHourStart, { hours: 17 }).toISOString(),
          timeZone: 'UTC'
        }
      };
  
      try {
        // POST /me/events
        await graphClient
          .api('/me/events')
          .post(socialHour);
  
        // Refresh the view
        this.render();
      } catch (error) {
        this.renderGraphError(error);
      }
    }
    // </addSocialToCalendarSnippet>

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
