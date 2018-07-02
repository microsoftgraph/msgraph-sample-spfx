import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphEventsListWebPartStrings';
import GraphEventsList from './components/GraphEventsList';
import { IGraphEventsListProps } from './components/IGraphEventsListProps';

import { MSGraphClient } from '@microsoft/sp-client-preview';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphEventsListWebPartProps {
  description: string;
}

export default class GraphEventsListWebPart extends BaseClientSideWebPart<IGraphEventsListWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGraphEventsListProps> = React.createElement(
      GraphEventsList,
      {
        graphClient: this.context.serviceScope.consume(MSGraphClient.serviceKey)
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
