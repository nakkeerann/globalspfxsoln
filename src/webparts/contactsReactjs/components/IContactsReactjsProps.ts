import { WebPartContext } from '@microsoft/sp-webpart-base';

import * as microsoftTeams from '@microsoft/teams-js';
import { ClientMode } from './ClientMode';
export interface IContactsReactjsProps {
  clientMode: ClientMode;
  teamsContext: microsoftTeams.Context;
  description: string;
  context: WebPartContext;
}
