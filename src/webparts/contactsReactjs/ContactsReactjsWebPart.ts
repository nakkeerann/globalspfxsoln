import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';

import * as strings from 'ContactsReactjsWebPartStrings';
import ContactsReactjs from './components/ContactsReactjs';
import { IContactsReactjsProps } from './components/IContactsReactjsProps';
import * as microsoftTeams from '@microsoft/teams-js';
import {ClientMode} from './components/ClientMode';
export interface IContactsReactjsWebPartProps {
  description: string;
  clientMode: ClientMode;
}

export default class ContactsReactjsWebPart extends BaseClientSideWebPart<IContactsReactjsWebPartProps> {
  private _teamsContext: microsoftTeams.Context;
  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    console.log(this.context.microsoftTeams);
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }
  
  public render(): void {
    const element: React.ReactElement<IContactsReactjsProps > = React.createElement(
      ContactsReactjs,
      {
        clientMode: this.properties.clientMode,
        teamsContext: this._teamsContext,
        context: this.context,
        description: this.properties.description
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
                }),
                PropertyPaneChoiceGroup('clientMode', {
                  label: strings.ClientModeLabel,
                  options: [
                    { key: ClientMode.aad, text: "AadHttpClient"},
                    { key: ClientMode.graph, text: "MSGraphClient"},
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
