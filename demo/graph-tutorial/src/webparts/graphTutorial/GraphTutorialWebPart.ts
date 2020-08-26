import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { startOfWeek, endOfWeek } from 'date-fns';
import { Providers, SharePointProvider, MgtAgenda } from '@microsoft/mgt';

import styles from './GraphTutorialWebPart.module.scss';
import * as strings from 'GraphTutorialWebPartStrings';

export interface IGraphTutorialWebPartProps {
  description: string;
}

export default class GraphTutorialWebPart extends BaseClientSideWebPart<IGraphTutorialWebPartProps> {

  // <onInitSnippet>
  protected async onInit() {
    // Set the toolkit's global auth provider to
    // SharePoint provider, allowing it to use the Graph
    // access token
    Providers.globalProvider =
      new SharePointProvider(this.context);
  }
  // </onInitSnippet>

  // <renderSnippet>
  /*
  public render(): void {
    this.context.msGraphClientFactory
      .getClient()
      .then((graphClient: MSGraphClient)=> {
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
          .get((error: any, events: any) => {
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
          });
      });
  } */
  // </renderSnippet>

  // <alternateRenderSnippet>
  public render(): void {
    // Get current date
    const now = new Date();
    // Get the start of the week based on current date
    const weekStart = startOfWeek(now);

    this.domElement.innerHTML = `
    <div class="${ styles.graphTutorial }">
      <div class="${ styles.container }">
        <div class="${ styles.row }">
          <div class="${ styles.column }">
            <mgt-agenda
              date="${weekStart.toISOString()}"
              days="7"
              group-by-day></mgt-agenda>
          </div>
        </div>
      </div>
    </div>`;
  }
  // </alternateRenderSnippet>

  // <renderCalendarViewSnippet>
  private renderCalendarView(events: MicrosoftGraph.Event[]) : void {
    const viewContainer = this.domElement.querySelector('#calendarView');

    // Create an agenda component from the toolkit
    let agenda = new MgtAgenda();
    // Set the events
    agenda.events = events;
    // Group events by day
    agenda.groupByDay = true;

    viewContainer.appendChild(agenda);
  }
  // </renderCalendarViewSnippet>

  // <renderGraphErrorSnippet>
  private renderGraphError(error: any): void {
    const viewContainer = this.domElement.querySelector('#calendarView');

    // Basic error display
    viewContainer.innerHTML = `
    <h2 class="${ styles.title }">Error</h2>
    <code style="word-break: break-all;">${JSON.stringify(error, null, 2)}</code>`;
  }
  // </renderGraphErrorSnippet>

  // @ts-ignore
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
