import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './components/GraphWebPart.module.scss';
import * as strings from 'GraphWebPartWebPartStrings';
import GraphWebPart from './components/GraphWebPart';
import { IGraphWebPartProps } from './components/IGraphWebPartProps';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IGraphWebPartWebPartProps {
  description: string;
}

export default class GraphWebPartWebPart extends BaseClientSideWebPart<IGraphWebPartWebPartProps> {

  private _renderEmailList(messages: MicrosoftGraph.Message[]): void {
    let html: string = '';
    for (let index = 0; index < messages.length; index++) {
      html += `<p class="${styles.description}">Email ${index + 1} - ${escape(messages[index].subject)}</p>`;
    }
  
    // Add the emails to the placeholder
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  public render(): void {
    this.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {
      // get information about the current user from the Microsoft Graph
      client
      .api('/me/messages')
      .top(5)
      .orderby("receivedDateTime desc")
      .get((error, messages: any, rawResponse?: any) => {
  
        this.domElement.innerHTML = `
        <div class="${ styles.graphWebPart}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Use Microsoft Graph in SharePoint Framework.</p>
              <div id="spListContainer" />
            </div>
          </div>
        </div>
        </div>`;
  
        // List the latest emails based on what we got from the Graph
        this._renderEmailList(messages.value);
  
      });
    });
  
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
