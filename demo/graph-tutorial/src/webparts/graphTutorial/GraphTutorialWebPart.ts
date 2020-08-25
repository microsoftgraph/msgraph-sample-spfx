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

import styles from './GraphTutorialWebPart.module.scss';
import * as strings from 'GraphTutorialWebPartStrings';

export interface IGraphTutorialWebPartProps {
  description: string;
}

export default class GraphTutorialWebPart extends BaseClientSideWebPart<IGraphTutorialWebPartProps> {

  // <renderSnippet>
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
  }
  // </renderSnippet>

  // <renderCalendarViewSnippet>
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
