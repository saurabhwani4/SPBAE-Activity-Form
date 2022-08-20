import { WebPartContext } from '@microsoft/sp-webpart-base';  
import { SPHttpClient } from '@microsoft/sp-http';  

export interface IWwspbaeReactProps {
  context: WebPartContext; 
  listName: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
}


