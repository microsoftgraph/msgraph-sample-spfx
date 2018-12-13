import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphTasksWebPartStrings';
import GraphTasks from './components/GraphTasks';
import { IGraphTasksProps } from './components/IGraphTasksProps';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphTasksWebPartProps {
  description: string;
}

export default class GraphTasksWebPart extends BaseClientSideWebPart<IGraphTasksWebPartProps> {

  public render(): void {
    this.context.msGraphClientFactory.getClient()
    .then((client: MSGraphClient): void => {
      const element: React.ReactElement<IGraphTasksProps> = React.createElement(
        GraphTasks,
        {
          graphClient: client
        }
      );

      ReactDom.render(element, this.domElement);
    });
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
