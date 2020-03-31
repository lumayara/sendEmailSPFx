import {MSGraphClientFactory} from '@microsoft/sp-http';
export interface ISendEmailProps {
  userEmail: string;
  graph: MSGraphClientFactory;
}
