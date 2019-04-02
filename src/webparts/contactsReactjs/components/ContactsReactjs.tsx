import * as React from 'react';
import styles from './ContactsReactjs.module.scss';
import { IContactsReactjsProps } from './IContactsReactjsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as microsoftTeams from '@microsoft/teams-js';
import { MSGraphClient } from '@microsoft/sp-http';
import * as strings from 'ContactsReactjsWebPartStrings';
import { ClientMode } from './ClientMode';
import { IUserItem } from './IUserItem';
import IContactsReactjsState from './IContactsReactjsState';
import { DetailsList, DetailsListLayoutMode, autobind,
  CheckboxVisibility, SelectionMode,
  TextField} from 'office-ui-fabric-react';
export default class ContactsReactjs extends React.Component<IContactsReactjsProps, IContactsReactjsState> {
  constructor(props: IContactsReactjsProps, state: IContactsReactjsState){
    super(props);
    this.state={
      users:[],
      searchFor:''
    };
  }
  public render(): React.ReactElement<IContactsReactjsProps> {
    let title: string = '';
    let subTitle: string = '';
    let siteTabTitle: string = '';
    // Configure the columns for the DetailsList component
let _usersListColumns = [
  {
    key: 'displayName',
    name: 'Display name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'mail',
    name: 'Mail',
    fieldName: 'mail',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'userPrincipalName',
    name: 'User Principal Name',
    fieldName: 'userPrincipalName',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
];
    if (this.props.teamsContext) {
      // We have teams context for the web part
      title = "Welcome to Teams!";
      subTitle = "Customize for your need";
      siteTabTitle = "We are in the context of following Team: " + this.props.teamsContext.teamName;
      this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // use MSGraphClient here
      });
    }
    else
    {
      // We are rendered in normal SharePoint context
      title = "Welcome to SharePoint!";
      subTitle = "Customize SharePoint experiences using Web Parts.";
      siteTabTitle = "We are in the context of following site: " + this.props.context.pageContext.web.title;
    }
  this._search();
    return (
      <div className={ styles.contactsReactjs }>
        <div className={ styles.container }>
        <TextField 
                    label={ strings.SearchFor } 
                    required={ true } 
                    value={ this.state.searchFor }
                    onChanged={ this._onSearchForChanged }
                    onGetErrorMessage={ this._getSearchForErrorMessage }
                  />
              {
                
                (this.state.users != null && this.state.users.length > 0) ?
                  <p>
                    
                  <DetailsList
                      items={ this.state.users }
                      columns={ _usersListColumns }
                      setKey='set'
                      checkboxVisibility={ CheckboxVisibility.hidden }
                      selectionMode={ SelectionMode.none }
                      layoutMode={ DetailsListLayoutMode.fixedColumns }
                      compact={ true }
                  />
                </p>
                : null
              }
            </div>
          
      </div>
    );
  }
  @autobind
    private _onSearchForChanged(newValue: string): void {

      // Update the component state accordingly to the current user's input
      this.setState({
        searchFor: newValue,
      });
    }

    private _getSearchForErrorMessage(value: string): string {
      // The search for text cannot contain spaces
      return (value == null || value.length == 0 || value.indexOf(" ") < 0)
        ? ''
        : `${strings.SearchForValidationErrorMessage}`;
    }
  private _search(): void {

    console.log(this.props.clientMode);

    // Based on the clientMode value search users
    switch (this.props.clientMode)
    {
      case ClientMode.aad:
        break;
      case ClientMode.graph:
        this._searchWithGraph();
        break;
    }
  }
  private _searchWithGraph(): void {

    // Log the current operation
    console.log("Using _searchWithGraph() method");

    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
        client
          .api("users")
          .version("v1.0")
          .select("displayName,mail,userPrincipalName")
          .filter(`(givenName eq '${escape(this.state.searchFor)}') or (surname eq '${escape(this.state.searchFor)}') or (displayName eq '${escape(this.state.searchFor)}')`)
          .get((err, res) => {  

            if (err) {
              console.error(err);
              return;
            }

            // Prepare the output array
            var users: Array<IUserItem> = new Array<IUserItem>();

            // Map the JSON response to the output array
            res.value.map((item: any) => {
              users.push( { 
                displayName: item.displayName,
                mail: item.mail,
                userPrincipalName: item.userPrincipalName,
              });
            });

            // Update the component state accordingly to the result
            this.setState(
              {
                users: users,
              }
            );
          });
      });
  }
}
