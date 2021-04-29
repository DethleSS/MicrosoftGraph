import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GraphWebPartWebPartStrings';
import { GraphWebPart } from './components/GraphWebPart';
import { IGraphWebPartProps } from './components/IGraphWebPartProps';

export interface IGraphWebPartWebPartProps {
  description: string;
  contextGraphApi: any;
  context: any;
}

export default class GraphWebPartWebPart extends BaseClientSideWebPart<IGraphWebPartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGraphWebPartProps> = React.createElement(
      GraphWebPart,
      {
        description: this.properties.description, 
        contextGraphApi: this.context.msGraphClientFactory,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
