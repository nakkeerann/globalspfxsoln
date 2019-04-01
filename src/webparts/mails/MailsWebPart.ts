import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MailsWebPart.module.scss';
import * as strings from 'MailsWebPartStrings';
import * as microsoftTeams from '@microsoft/teams-js';
export interface IMailsWebPartProps {
  description: string;
}

export default class MailsWebPart extends BaseClientSideWebPart<IMailsWebPartProps> {
  private _teamsContext: microsoftTeams.Context;
  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
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

    let title: string = '';
    let subTitle: string = '';
    let siteTabTitle: string = '';
  
    if (this._teamsContext) {
      // We have teams context for the web part
      title = "Welcome to Teams!";
      subTitle = "Customize for your need";
      siteTabTitle = "We are in the context of following Team: " + this._teamsContext.teamName;
    }
    else
    {
      // We are rendered in normal SharePoint context
      title = "Welcome to SharePoint!";
      subTitle = "Customize SharePoint experiences using Web Parts.";
      siteTabTitle = "We are in the context of following site: " + this.context.pageContext.web.title;
    }
  
    this.domElement.innerHTML = `
      <div>
              <span class="${ styles.title }">${title}</span>
              <p class="${ styles.subTitle }">${subTitle}</p>
              <p class="${ styles.description }">${siteTabTitle}</p>
          
      </div>`;
  
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
