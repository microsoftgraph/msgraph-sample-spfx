import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { AadTokenProvider } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GraphTutorialWebPart.module.scss';
import * as strings from 'GraphTutorialWebPartStrings';

export interface IGraphTutorialWebPartProps {
  description: string;
}

export default class GraphTutorialWebPart extends BaseClientSideWebPart<IGraphTutorialWebPartProps> {

  public render(): void {
    this.context.aadTokenProviderFactory
      .getTokenProvider()
      .then((provider: AadTokenProvider)=> {
        provider
          .getToken('https://graph.microsoft.com')
          .then((token: string) => {
            this.domElement.innerHTML = `
            <div class="${ styles.graphTutorial }">
              <div class="${ styles.container }">
                <div class="${ styles.row }">
                  <div class="${ styles.column }">
                    <span class="${ styles.title }">Welcome to SharePoint!</span>
                    <p><code style="word-break: break-all;">${ token }</code></p>
                  </div>
                </div>
              </div>
            </div>`;
          });
      });
  }

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
