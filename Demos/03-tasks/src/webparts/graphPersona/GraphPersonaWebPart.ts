import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphPersonaWebPartStrings';
import GraphPersona from './components/GraphPersona';
import { IGraphPersonaProps } from './components/IGraphPersonaProps';

import { MSGraphClient } from '@microsoft/sp-client-preview';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphPersonaWebPartProps {
  description: string;
}

export default class GraphPersonaWebPart extends BaseClientSideWebPart<IGraphPersonaWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGraphPersonaProps > = React.createElement(
      GraphPersona,
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
