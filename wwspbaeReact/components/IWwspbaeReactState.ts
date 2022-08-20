import { IListItem } from './IListItem';  

  
export interface IWwspbaeReactState {  
  status: string;  
  items: IListItem[];  
  hidden:boolean;
  users: any [];
  readUsers: any [];
  currentUser: number;
  latestListItem: number;
  currentListItem: number;
  defaultUser: string [];
  readSubmitterID: number;
  dropdown_items_account: any [];
  validation_msg: string;
  newlyAddedItem: number;
  dropdown_items_location: any[];
  dropdown_items_location_specific: any[];
  readLocation: any[];
  account_selected: any[];
  date_for_title: string;
  email_alias_title: string;
  NotifyTo_users: any [];
  readNotifyTo: any[];
  defaultNotifyTo: string [];
  dropdown_items_email: any [];
  emailAssociated: string;
  loading: boolean;
  attendee_suggestions: any[];
  nameValue: string;
  nameSuggestions: any[];
  emailValue: string;
  emailSuggestions: any[];
  roleSuggestions: any[];
  roleValue: string;
}  