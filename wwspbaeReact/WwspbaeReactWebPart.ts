import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { escape } from '@microsoft/sp-lodash-subset';  
import * as strings from 'WwspbaeReactWebPartStrings';
import WwspbaeReact from './components/WwspbaeReact';
import { IWwspbaeReactProps } from './components/IWwspbaeReactProps';

import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";

export interface IWwspbaeReactWebPartProps {
  context: WebPartContext;
  listName: string;
}

export default class WwspbaeReactWebPart extends BaseClientSideWebPart<IWwspbaeReactWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IWwspbaeReactProps> = React.createElement(
      WwspbaeReact,
      {
        context: this.context,
        listName: this.properties.listName,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,  
      }
      
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
    sp.setup({
    spfxContext: this.context
    });
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
                PropertyPaneTextField('ListName', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
