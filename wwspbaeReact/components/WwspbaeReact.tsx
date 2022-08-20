import * as React from 'react';
import { IWwspbaeReactProps } from './IWwspbaeReactProps';
import { IWwspbaeReactState } from './IWwspbaeReactState';
import { escape, groupBy, isEmpty } from '@microsoft/sp-lodash-subset';
import { IDropdownOption } from "office-ui-fabric-react";
import {PeoplePicker, PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import '../styles.css';
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
import * as $ from 'jquery';
import { List, sp, View } from "@pnp/sp";
import { CurrentUser } from '@pnp/sp/src/siteusers';
import { IListItem } from './IListItem';  
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'; 
import { peoplePicker } from 'office-ui-fabric-react/lib/components/FloatingPicker/PeoplePicker/PeoplePicker.scss';
import TextField from "@material-ui/core/TextField";
import Autocomplete from "@material-ui/lab/Autocomplete";
import { Multiselect } from "multiselect-react-dropdown";
import { data, type } from 'jquery';
import { css } from "@emotion/react";
import ClipLoader from "react-spinners/ClipLoader";
import Autosuggest from 'react-autosuggest';
import autoBind from 'react-autobind';


// https://developer.mozilla.org/en/docs/Web/JavaScript/Guide/Regular_Expressions#Using_Special_Characters
function escapeRegexCharacters(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function getSuggestionname(suggestion) {
  return suggestion.name;
}

function getSuggestionEmail(suggestion) {
  return suggestion.email;
}

function getSuggestionRole(suggestion) {
  return suggestion.role;
}

function renderSuggestion(suggestion) {
  return (
    <span>{suggestion.name} - {suggestion.email}</span>
  );
}

// function to split the value in Accounts column and autofill Location
function getLocation() {
  var e = $('#dropdown').val().toString();
  var array = e.split('/');
  var acc = array[1];
  $("#loc").val(acc);
}


function validationMessage(a: string){
  var $messageDiv = $('#validation'); // get the reference of the div
  $messageDiv.show().html(a); // show and set the message
  setTimeout(function(){ $messageDiv.hide().html('');}, 3000);
};


// Check if the inputs of Attendee's details are empty before adding the row 



// Returns checked values from the checkbox in an array that can then be used to send to SP choice column

function checkbox_array()
{
	var all_location_id = document.querySelectorAll('input[name="location[]"]:checked');
	var aIds = [];
	for(var x = 0, l = all_location_id.length; x < l;  x++)
	{
	    aIds.push($(all_location_id[x]).val());
	}
	return aIds;
}

/* function checkbox()
{
	var all_location_id = document.querySelectorAll('input[name="location[]"]:checked');
	var aIds = [];
	for(var x = 0, l = all_location_id.length; x < l;  x++)
	{
	    aIds.push($(all_location_id[x]).val());
	}
	var str = aIds.join(', ');
	return str;
}*/

// Return the input values from the Attendee's table. comma separate the row values, semi colon separate the column values 

function showTableData(){
  var TableData = '';
   $('#tbody_saurabh > tr').each(function(row, tr){
  TableData = TableData 
      + $(tr).find('td:eq(1)').text() + ','  // Task No.
      + $(tr).find('td:eq(2)').text() + ','  // Date
      + $(tr).find('td:eq(3)').text() + ''  // Description
      + ';';
});
   return(TableData);
}

// Read the values into table from sp list

function readTableData(a: string){
  var b = a.split(';');
  
  var i: number;
  for (i=0; i<b.length-1; ++i){
    var c = b[i].split(',');
    var name = c[0];
    var email = c[1];
    var title = c[2];

    var markup = "<tr><td><input type='checkbox'  name='record'></td><td>" + name + "</td><td>" + email + "</td><td>" + title + "</td></tr>";
    $("#tbody_saurabh").append(markup);

  }
 }

// Read Checkbox values

function readCheckBox(a){
  var i: number;
  var j: number;
  let checkboxes = $("#myChecks :input");
  for (j=0;j<a.length;++j){
    for (i=0;i<4;++i){
    if ($(checkboxes[i]).val() == a[j]){
      $(checkboxes[i]).prop('checked',true);
    }
  }
  }
}

function GetUrlParameter(sParam)
{
    var sPageURL = window.location.search.substring(1);
    var sURLVariables = sPageURL.split('&');
    for (var i = 0; i < sURLVariables.length; i++)
    {
        var sParameterName = sURLVariables[i].split('=');
        if (sParameterName[0] == sParam)
        {
            return sParameterName[1];
        }
    }
}



var vdropdown_items_email = [];
var vdropdown_items_location = [];
var vdropdown_items_attendees = [];

export default class WwspbaeReact extends React.Component<IWwspbaeReactProps, IWwspbaeReactState, {}> {
  multiselectRef: any;

  constructor(props: IWwspbaeReactProps, state: IWwspbaeReactState){
    super(props);
    this.multiselectRef = React.createRef();
    autoBind(this);
    this.state = {
      status: '',
      items: [],
      hidden: false,
      users: [],
      readUsers: [],
      currentUser: 0,
      latestListItem: 0,
      currentListItem: 0,
      defaultUser: [],
      readSubmitterID: 0,
      dropdown_items_account: [],
      validation_msg: '',
      newlyAddedItem: 0,
      dropdown_items_location: [],
      dropdown_items_location_specific: [],
      readLocation: [],
      account_selected: [],
      date_for_title: '',
      email_alias_title: '',
      NotifyTo_users: [],
      readNotifyTo: [],
      defaultNotifyTo: [],
      dropdown_items_email: [],
      emailAssociated: '',
      loading: false,
      nameValue: '',
      nameSuggestions: [],
      emailValue: '',
      emailSuggestions: [],
      attendee_suggestions: [],
      roleValue: '',
      roleSuggestions: [],
    };

  }
  @autoBind
  public getSuggestions(value) {
    const escapedValue = escapeRegexCharacters(value.trim());
    const regex = new RegExp('^' + escapedValue, 'i');
    
    return this.state.attendee_suggestions.filter(user => regex.test(user.name) || regex.test(user.email) || regex.test(user.role));
  }

  @autoBind
  public onnameChange = async (event, { newValue }) => {
    this.setState({
      nameValue: newValue
    });
    
  };

  @autoBind
  public onEmailChange = async(event, { newValue }) => {
    this.setState({
      emailValue: newValue
    });
  };

  @autoBind
  public onRoleChange = async(event, { newValue }) => {
    this.setState({
      roleValue: newValue
    });
  };

  @autoBind
  public onnameSuggestionsFetchRequested = async({ value }) => {
    this.setState({
      nameSuggestions: this.getSuggestions(value)
    });
  };

  @autoBind
  public onnameSuggestionsClearRequested = () => {
    this.setState({
      nameSuggestions: []
    });
  };

  @autoBind
  public onnameSuggestionSelected = async (event, { suggestion }) => {
    await this.setState({
      nameValue: suggestion.name,
      emailValue: suggestion.email,
      roleValue: suggestion.role
    });
    console.log(this.state.nameValue);
  };

  @autoBind
  public onEmailSuggestionsFetchRequested = ({ value }) => {
    this.setState({
      emailSuggestions: this.getSuggestions(value)
    });
  };

  @autoBind
  public onEmailSuggestionsClearRequested = () => {
    this.setState({
      emailSuggestions: []
    });
  };

  @autoBind
  public onEmailSuggestionSelected = async(event, { suggestion }) => {
    await this.setState({
      nameValue: suggestion.name,
      emailValue: suggestion.email,
      roleValue: suggestion.role
    });
  };

  @autoBind
  public onRoleSuggestionsFetchRequested = ({ value }) => {
    this.setState({
      roleSuggestions: this.getSuggestions(value)
    });
  };

  @autoBind
  public onRoleSuggestionsClearRequested = () => {
    this.setState({
      roleSuggestions: []
    });
  };

  @autoBind
  public onRoleSuggestionSelected = async(event, { suggestion }) => {
    await this.setState({
      emailValue: suggestion.email,
      nameValue: suggestion.name,
      roleValue: suggestion.role
    });
  };

  @autoBind
  public checkEmptyInput(){

    /*var isEmpty = false,
        att_name = $("#att_name").val(),
        att_email = $("#att_email").val(),
        att_title = $("#att_title").val();
    */
    var isEmpty = false,
    att_name = this.state.nameValue,
    att_email = this.state.emailValue,
    att_title = this.state.roleValue;

    if(att_name === ""){
        
        validationMessage("Attendee's Name Connot Be Empty");
        isEmpty = true;
    }
    else if(att_email === ""){
        validationMessage("Attendee's Email Connot Be Empty");
        isEmpty = true;
    }
    else if(att_title === ""){
        validationMessage("Attendee's Title Connot Be Empty");
        isEmpty = true;
    }
    return isEmpty;
}

// Email Validation

@autoBind
  public validateEmail($email: any) {
    var emailReg = /^([\w-\.]+@([\w-]+\.)+[\w-]{2,4})?$/;
    return emailReg.test( $email );
  }

  @autoBind
  // Add a row to the table
  public addRow(){
    /*if(!checkEmptyInput()){
      var name = $("#att_name").val();
      var email = $("#att_email").val();
      var title = $("#att_title").val();
      var markup = "<tr><td><input type='checkbox' name='record'></td><td>" + name + "</td><td>" + email + "</td><td>" + title + "</td></tr>";

      if( !validateEmail(email)) { 
        validationMessage("Attendee's Email is incorrect");
      }
      else{
        $("#tbody_saurabh").append(markup);
        $("#att_name").val("");
        $("#att_email").val("");
        $("#att_title").val("");
      }
    }*/
    if(!this.checkEmptyInput()){
      var name = this.state.nameValue;
      console.log(name);
      var email = this.state.emailValue;
      console.log(email);
      var title = this.state.roleValue;
      console.log(title);
      var markup = "<tr><td><input type='checkbox' name='record'></td><td>" + name + "</td><td>" + email + "</td><td>" + title + "</td></tr>";

      if( !this.validateEmail(email)) { 
        validationMessage("Attendee's Email is incorrect");
      }
      else{
        $("#tbody_saurabh").append(markup);
        this.setState({
          nameValue: "",
          emailValue: "",
          roleValue: ""
        });
      }
  }
}
  

  // delete selected row from the table

  @autoBind
  public deleteRow(){
    $("#tbody_saurabh").find('input[name="record"]').each(function(){
      if($(this).is(":checked")){
          $(this).closest("tr").remove();
      }
  });
  }

  @autoBind
  public loaderSpinner(){
    $("#container").css({"background": "#D6DBDF", "opacity": "0.6"});
    $("input").css({"background": "#D6DBDF", "opacity": "0.6"});   
  }
  @autoBind
  public removeLoaderSpinner(){
    $("#container").css({"background": "", "opacity": ""});
    $("input").css({"background": "", "opacity": ""});
  }

  @autoBind
  public async componentDidMount(){
    this.date();
    //console.log('Default User Before: '+this.state.defaultUser)
    this.getLatestItemId().then((itemId: number) => {  
      this.setState({
        latestListItem: itemId,
        currentListItem: parseInt(GetUrlParameter('ID')),
      });
      if (this.state.currentListItem <= this.state.latestListItem){
        this.readItem();
        this.getSPData_editform();
      }
      else{
        this.getSPData_newform();
      }
    });

    $('#validation').hide();
    $('#notification').hide();
    $('#notification_label').hide();

    // Get all the Account names from a sharepoint list

    await sp.web.lists.getByTitle("Activity Report - Account and Location").items.select('Title','Location').get().then(function(data){
      
      //List uniquevalues = data.ToList<SPListItem>();
      for(var k in data){
        vdropdown_items_location.push({Company:data[k].Title, location:data[k].Location});
      }      
    });

    await sp.web.lists.getByTitle("Activity Report - Accounts").items.select('Title','EmailAssociated').get().then(function(data){
      for(var k in data){
        vdropdown_items_email.push({Company:data[k].Title, Email:data[k].EmailAssociated});
      }        
    });

    let result = vdropdown_items_email.map(a => a.Company);
    //console.log(result);

    this.setState({
      dropdown_items_account: result,
      dropdown_items_location: vdropdown_items_location,
      dropdown_items_email: vdropdown_items_email
    });
    this._getdefaultPeoplePickerItems(); 
    this._getdefaultNotifyToItems();

  }

  @autoBind
  private async getSPData_newform(){
    await sp.web.currentUser.get().then((response: CurrentUser) => {
      $('#submitter').val(response["Title"]);
      var email_alais = response["Email"].split('@')[0];
      this.setState({
        currentUser: response["Id"],
        email_alias_title: email_alais
      });
    });
    } 

  @autoBind
  private async getSPData_editform(){
    await sp.web.currentUser.get().then((response: CurrentUser) => {
      var email_alais = response["Email"].split('@')[0];
      this.setState({
        currentUser: response["Id"],
        email_alias_title: email_alais
      });
    });
    }  
  
  @autoBind
  private async Submitter(){
    await sp.web.currentUser.get().then((response: CurrentUser) => {
      let id: number = response["Id"];
    });
    return 0;
    } 

  @autoBind
  private async readSubmitter(x: any) {  
    await this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/getuserbyid(${x})`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'odata-version': ''  
      }  
    }) 
    .then((response: SPHttpClientResponse): Promise<any> => {  
      return response.json();  
    })  
    .then((item: any): void => {
      $("#submitter").val(item["Title"]);
      }
    );  
  }

  // function to get today's date - returns a string

@autoBind
public date() {
  var date = new Date();

	var day = (date.getDate()).toString();
	var month = (date.getMonth() + 1).toString();
	var year = (date.getFullYear()).toString();

	if (Number(month) < 10) month = "0" + month;
	if (Number(day) < 10) day = "0" + day;

	var today = year + "-" + month + "-" + day;       
  var date_for_ttl = year + month + day;

  this.setState({
    date_for_title: date_for_ttl
  });  

	$('#date').val(today);
}
  
  @autoBind
  private async _getPeoplePickerItems(items: any[]) {
    var i: number;
    var userIDs = [];
    for (i=0;i<items.length;++i){
      console.log(items[i]);
      var a = items[i].secondaryText;
      var result = await sp.web.ensureUser(a);
      userIDs.push(result.data.Id);
    }
    console.log(userIDs);  
    this.setState({
      users: userIDs,
    });
  }
  
  @autoBind
  private async _getdefaultPeoplePickerItems() {
    var items: string[];
    //console.log(this.state.readUsers);
    await this.readPeopleItem(this.state.readUsers);
    items = this.state.defaultUser;
    //console.log("DefaultUser: "+items);
    var i: number;
    var userIDs = [];
    for (i=0;i<items.length;++i){
      var result = await sp.web.ensureUser(items[i]);
      userIDs.push(result.data.Id);
    }
    //console.log("UserIds: "+userIDs);
    this.setState({
      users: userIDs,
    });
  }

  @autoBind
  private async _getdefaultNotifyToItems() {
    var items: string[];
    //console.log(this.state.readUsers);
    await this.readNotifyToItem(this.state.readNotifyTo);
    items = this.state.defaultNotifyTo;
    //console.log("DefaultUser: "+items);
    var i: number;
    var userIDs = [];
    for (i=0;i<items.length;++i){
      var result = await sp.web.ensureUser(items[i]);
      userIDs.push(result.data.Id);
    }
    //console.log("UserIds: "+userIDs);
    this.setState({
      NotifyTo_users: userIDs,
    });
  }

  @autoBind
  private async _getNotifyToPeople(items: any[]) {
    var i: number;
    var userIDs = [];
    for (i=0;i<items.length;++i){
      console.log(items[i]);
      var a = items[i].secondaryText;
      var result = await sp.web.ensureUser(a);
      userIDs.push(result.data.Id);
    }
    console.log(userIDs);  
    this.setState({
      NotifyTo_users: userIDs,
    });
  }

  @autoBind
  public getTitle(){
    var d = new Date();
    var hrs = (d.getHours()).toString();
    var min = (d.getMinutes()).toString();
    var title = this.state.account_selected.toString() + '_' + this.state.date_for_title + hrs + min +'_' + this.state.email_alias_title;
    return title;
  }

  @autoBind
  private readLocation(x:string){
    var a: any[] = x.split(',');
    this.setState({
      readLocation: a,
    });
  }

  @autoBind
  public async fetchLocation(selectedList, selectedItem){
    await sp.web.lists.getByTitle("Attendee Details").items.select('Title','Title0','Email').get().then(function(data){
      for(var k in data){
        vdropdown_items_attendees.push({name:data[k].Title, email:data[k].Email});
      }        
    });

    let attendees = vdropdown_items_attendees
    console.log(attendees);

    await this.setState({
      readLocation: selectedList,
      attendee_suggestions: vdropdown_items_attendees,
    });
  }

  @autoBind
  public joinLocations(){
    var locations_string: string = this.state.readLocation.join(',');
    return locations_string;
  }

  @autoBind
  public async sort_location_based_on_selected_company(selectedList, selectedItem){
    //Making the sugestions null on change of account
    var vdropdown_items_attendees = [];

    var a = groupBy(this.state.dropdown_items_location, "Company")[selectedItem];
    let result = a.map(b => b.location);

    var x = groupBy(this.state.dropdown_items_email, "Company")[selectedItem];
    let result_email = x.map(y => y.Email);
    //console.log(result_email[0]);
    this.resetValues();

    await this.setState({
      dropdown_items_location_specific: result,
      account_selected: selectedList,
      emailAssociated: result_email[0],
      readLocation: []
    });

    var $messageDiv = $('#notification'); // get the reference of the div
    $messageDiv.show().html(result_email[0]); // show and set the message

    var $messageDiv_ = $('#notification_label'); // get the reference of the div
    $messageDiv_.show().html('The following recipients /email aliases will be notified:'); // show and set the message
    
    //Generate Sugestions based on account selected
    await sp.web.lists.getByTitle("Attendee Details").items.select('Title','Title0','Email').filter("AccountAssociated eq '" + this.state.account_selected[0] + "' ").getAll().then(function(data){
      for(var k in data){
        vdropdown_items_attendees.push({name:data[k].Title, email:data[k].Email, role: data[k].Title0});
      }        
    });

    let attendees = vdropdown_items_attendees
    console.log(attendees);

    await this.setState({
      attendee_suggestions: vdropdown_items_attendees,
    });

  }


  
  @autoBind
  private async createAttendeeItem(ID: number, name: string, email: string, title: string){  
  
      const body: string = JSON.stringify({  
        'Master_ID': ID,
        'Title': name, //Name
        'Email': email,
        'Title0': title, //Title
        'AccountAssociated': this.state.account_selected[0].toString()
      });  
      
      await this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('Attendee Details')/items`,  
      SPHttpClient.configurations.v1,  
      {  
        headers: {  
          'Accept': 'application/json;odata=nometadata',  
          'Content-type': 'application/json;odata=nometadata',  
          'odata-version': ''  
        },  
        body: body  
      })  
      .then((response: SPHttpClientResponse): Promise<IListItem> => {  
        return response.json();  
      });  
    }
  
    
  resetValues() {
    // By calling the belowe method will reset the selected values programatically
    this.multiselectRef.current.resetSelectedValues();
  };

  @autoBind
  private async createItem(x:string){  
    
    if (this.state.currentListItem <= this.state.latestListItem){
      var SubmitterId = this.state.readSubmitterID;
    }
    else{
      var SubmitterId = this.state.currentUser;
    }

    var submitter_name = $("#submitter").val();
    var date = $("#date").val();
    var acc_name = this.state.account_selected[0];
    var loc_name = this.joinLocations();
    var other_atte = this.state.users;
    var check = checkbox_array();
    var ex_sum = $("#exc_sum").val();
    var detail_sum = $("#det_sum").val();
    var att_details = showTableData();
    var title = this.getTitle();
    var notifyto = this.state.NotifyTo_users;
    var emailAssciated = this.state.emailAssociated;
  
    var all_data = {title: title, submitterId: SubmitterId, submitter: submitter_name, date: date, account_name: acc_name, location_name: loc_name, other_cadence_attendees: other_atte, purpose: check, executive: ex_sum, detailed: detail_sum, attendees_deails: att_details, notifytoUsers: notifyto, emailAssciated: emailAssciated};
    
    if (x === 'Saved'){

      const body: string = JSON.stringify({  
        'Title': all_data.title,
        'Location': all_data.location_name,
        'Attendees': all_data.attendees_deails,
        'Status': x,
        'AccountName': all_data.account_name,
        'ExecutiveSummary': all_data.executive,
        'DetailedSummary': all_data.detailed,
        'DateofVisit': all_data.date,
        'Purposeofthemeeting': all_data.purpose,
        'SubmitterId': all_data.submitterId,
        'OtherCadenceAttendeesId': all_data.other_cadence_attendees,
        'AuthorId': all_data.submitterId,
        'NotifyToId': all_data.notifytoUsers,
        'EmailAssociated': all_data.emailAssciated
      });

      if (this.state.currentListItem <= this.state.latestListItem){
        await this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('Activity Report')/items(${this.state.currentListItem})?$select=Title,SubmitterId,Location,DateofVisit,AccountName,ExecutiveSummary,DetailedSummary, Purposeofthemeeting, Attendees`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': '',  
            'IF-MATCH': '*',  
            'X-HTTP-Method': 'MERGE'  
          },  
          body: body  
        });
      }
      else{
        await this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('Activity Report')/items`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': ''  
          },  
          body: body  
        })  
        .then((response: SPHttpClientResponse): Promise<IListItem> => {  
          return response.json();  
        });
        
      }
      this.closeForm();
    }
    else{
    if(this.state.account_selected.length === 0){
      validationMessage('Please select an Account');
    }
    else if (this.state.readLocation.length === 0){
      validationMessage('Please select a Site');
    }
    else if (check.length === 0){
      validationMessage("Please select a purpose of the meeting");
    }
    else if (all_data.executive === ''){
      validationMessage('Please add executive summary');
    }
    else if (all_data.detailed === ''){
     validationMessage('Please add detailed summary');
    }
    else if (all_data.attendees_deails === ''){
      validationMessage("Please add attendee's information");
    }
    else{
      const body: string = JSON.stringify({  
        'Title': all_data.title,
        'Location': all_data.location_name,
        'Attendees': all_data.attendees_deails,
        'Status': x,
        'AccountName': all_data.account_name,
        'ExecutiveSummary': all_data.executive,
        'DetailedSummary': all_data.detailed,
        'DateofVisit': all_data.date,
        'Purposeofthemeeting': all_data.purpose,
        'SubmitterId': all_data.submitterId,
        'OtherCadenceAttendeesId': all_data.other_cadence_attendees,
        'AuthorId': all_data.submitterId,
        'NotifyToId': all_data.notifytoUsers,
        'EmailAssociated': all_data.emailAssciated
      });  
      
      if (this.state.currentListItem <= this.state.latestListItem){
        await this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('Activity Report')/items(${this.state.currentListItem})?$select=Title,SubmitterId,Location,DateofVisit,AccountName,ExecutiveSummary,DetailedSummary, Purposeofthemeeting, Attendees`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': '',  
            'IF-MATCH': '*',  
            'X-HTTP-Method': 'MERGE'  
          },  
          body: body  
        });
        this.setState({
          newlyAddedItem: this.state.currentListItem
        });
        this.attendeesDetails();
      }
      else{
        await this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('Activity Report')/items`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': ''  
          },  
          body: body  
        })  
        .then((response: SPHttpClientResponse): Promise<IListItem> => {  
          return response.json();  
        });  

        await this.getLatestItemId().then((itemId: number) => {
          this.setState({
            newlyAddedItem: itemId
          });
        }); 
        this.attendeesDetails();               
      }
      this.closeForm();
    }
    }
    
  }   

  @autoBind
  private attendeesDetails(){
    var allItems = [];
    $('#tbody_saurabh > tr').each(function(row, tr){
        var rowName = $(tr).find('td:eq(1)').text();     // Name
        var rowEmail = $(tr).find('td:eq(2)').text();   // Email
        var rowTitle = $(tr).find('td:eq(3)').text();  // Title
        var thisrow = [rowName, rowEmail, rowTitle];
        allItems.push(thisrow);
    }
    );

    var i: number;
    for(i=0; i<allItems.length; ++i){
       this.createAttendeeItem(this.state.newlyAddedItem, allItems[i][0], allItems[i][1], allItems[i][2]);
    }
  }


    
  
  @autoBind
  private async readItem(){  

    this.setState({
      loading: true,
    });

    //this.loaderSpinner();

    await this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Activity Report')/items(${this.state.currentListItem})?$select=Title,Location,DateofVisit,AccountName,ExecutiveSummary,DetailedSummary, Purposeofthemeeting, Attendees, OtherCadenceAttendeesId, Status, AuthorId, NotifyToId`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'odata-version': ''  
      }  
    }) 
    .then((response: SPHttpClientResponse): Promise<IListItem> => {  
      return response.json();  
    })  
    .then(async (item: IListItem) => { 
      this.readSubmitter(item.AuthorId);
      this.setState({
        readUsers: item.OtherCadenceAttendeesId,
        readSubmitterID: item.AuthorId,
        readNotifyTo: item.NotifyToId,
      });

      if((item.Status === "Saved")&&(this.state.currentUser == this.state.readSubmitterID)){ 
        this.setState({
          hidden: false,
        });
      }
      else{
        this.setState({
          hidden: true,
        });
      }
      
      $("#date").val(item.DateofVisit);
      $("#exc_sum").val(item.ExecutiveSummary);
      $("#det_sum").val(item.DetailedSummary);

      try{
          $("#loc").val(item.Location);
          var x = item.AccountName.concat('/',item.Location);
          $("#dropdown").val(x);
      }
      finally{
        try{
          readTableData(item.Attendees);
        }
        finally{
          try{
            readCheckBox(item.Purposeofthemeeting);
          }
          finally{
            try{
              var account_array = [item.AccountName]; 
              this.setState({
                account_selected: account_array,
              });
              this.readLocation(item.Location);

              var abc: any[];
              abc = groupBy(this.state.dropdown_items_email, "Company")[item.AccountName];
              let result_email = abc.map(y => y.Email);

              var $messageDiv = $('#notification'); // get the reference of the div
              $messageDiv.show().html(result_email[0]); // show and set the message

              var $messageDiv_ = $('#notification_label'); // get the reference of the div
              $messageDiv_.show().html('The following recipients /email aliases will be notified:'); // show and set the message
            }
            finally{
              try{
              await this._getdefaultPeoplePickerItems();
              }
              finally{
              try{
                await this._getdefaultNotifyToItems();
              }
              finally{

                //Generate Sugestions based on account selected
                var vdropdown_items_attendees = [];

                await sp.web.lists.getByTitle("Attendee Details").items.select('Title','Title0','Email').filter("AccountAssociated eq '" + this.state.account_selected[0] + "' ").getAll().then(function(data){
                  for(var k in data){
                    vdropdown_items_attendees.push({name:data[k].Title, email:data[k].Email, role: data[k].Title0});
                  }        
                });

                let attendees = vdropdown_items_attendees
                console.log(attendees);

                await this.setState({
                  attendee_suggestions: vdropdown_items_attendees,
                });

                this.setState({
                  loading: false,
                });
              }
              }
            }
          }
        }
      }
      //this.removeLoaderSpinner();
    },  
    );
  }
  
  /*@autobind
  private readOnlyItem(): void {  
    this.setState({
      hidden: true,
    });
  
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Activity Report')/items(107)?$select=Title,Location,DateofVisit,AccountName,ExecutiveSummary,DetailedSummary, Purposeofthemeeting, Attendees, OtherCadenceAttendeesId`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'odata-version': ''  
      }  
    }) 
    .then((response: SPHttpClientResponse): Promise<IListItem> => {  
      return response.json();  
    })  
    .then((item: IListItem): void => {  
      if(item.Status === "Submitted"){ 
        $("#date").val(item.DateofVisit);
        $("#loc").val(item.Location);
        $("#dropdown_account option:contains(" + item.Location + ")").attr('selected', 'selected');
        $("#exc_sum").val(item.ExecutiveSummary);
        $("#det_sum").val(item.DetailedSummary);
        readTableData(item.Attendees);
        readCheckBox(item.Purposeofthemeeting);
        this.setState({
          readUsers: item.OtherCadenceAttendeesId
        });
      }
    },  
    );
  }*/
  
  
  /*private updateItem(): void {   
    
      var submitter_name = $("#submitter").val();
      var SubmitterId = this.state.currentUser;
      var date = $("#date").val();
      var acc_name = $("#dropdown_account").val();
      var loc_name = $("#loc").val();;
      var other_atte = this.state.users;
      var check = checkbox_array();
      var ex_sum = $("#exc_sum").val();
      var detail_sum = $("#det_sum").val();
      var att_details = showTableData();
    
      
      var all_data = {submitterId: SubmitterId, submitter: submitter_name, date: date, account_name: acc_name, location_name: loc_name, other_cadence_attendees: other_atte, purpose: check, executive: ex_sum, detailed: detail_sum, attendees_deails: att_details};
      const body: string = JSON.stringify({  
        'Title': 'Title',
        'Location': all_data.location_name,
        'Attendees': all_data.attendees_deails,
        'Status': 'Saved',
        'AccountName': all_data.account_name,
        'ExecutiveSummary': all_data.executive,
        'DetailedSummary': all_data.detailed,
        'DateofVisit': all_data.date,
        'Purposeofthemeeting': all_data.purpose,
        'SubmitterId': all_data.submitterId,
        'OtherCadenceAttendeesId': all_data.other_cadence_attendees
      });  
  
      if (this.state.currentListItem <= this.state.latestListItem){
        this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('Activity Report')/items(${this.state.currentListItem})?$select=Title,Location,DateofVisit,AccountName,ExecutiveSummary,DetailedSummary, Purposeofthemeeting, Attendees`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': '',  
            'IF-MATCH': '*',  
            'X-HTTP-Method': 'MERGE'  
          },  
          body: body  
        })
      }
      else{
        this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('Activity Report')/items`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': ''  
          },  
          body: body  
        })  
        .then((response: SPHttpClientResponse): Promise<IListItem> => {  
          return response.json();  
        })
      }
        
  
  }*/
  @autoBind
  private getLatestItemId(): Promise<number> {  
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {  
      sp.web.lists.getByTitle('Activity Report')  
        .items.orderBy('Id', false).top(1).select('Id').get()  
        .then((items: { Id: number }[]): void => {  
          if (items.length === 0) {  
            resolve(-1);  
          }  
          else {  
            resolve(items[0].Id);  
          }  
        }, (error: any): void => {  
          reject(error);  
        });  
    });  
  }  
  
  
  @autoBind
  public async readPeopleItem(x: any) {  
    var userEmails: string [];
    userEmails = []; 
    var i: number;

    for (i=0;i<x.length;++i){
    await this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/getuserbyid(${x[i]})`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'odata-version': ''  
      }  
    }) 
    .then((response: SPHttpClientResponse): Promise<any> => {  
      return response.json();  
    })  
    .then((item: any): void => {
      userEmails.push(item["Email"]);
      //console.log('User Emails single: '+item["Email"])
      }
    );  
    }

    this.setState({
      defaultUser: userEmails,
    });
    //console.log('User Emails Total: '+userEmails)
    //console.log('Default User After: '+this.state.defaultUser)
    return(userEmails);
    
  }

  @autoBind
  public async readNotifyToItem(x: any) {  
    var userEmails: string [];
    userEmails = []; 
    var i: number;

    for (i=0;i<x.length;++i){
    await this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/getuserbyid(${x[i]})`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'odata-version': ''  
      }  
    }) 
    .then((response: SPHttpClientResponse): Promise<any> => {  
      return response.json();  
    })  
    .then((item: any): void => {
      userEmails.push(item["Email"]);
      //console.log('User Emails single: '+item["Email"])
      }
    );  
    }

    this.setState({
      defaultNotifyTo: userEmails,
    });
    //console.log('User Emails Total: '+userEmails)
    //console.log('Default User After: '+this.state.defaultUser)
    return(userEmails);
    
  }

  public closeForm(){
    window.location.href = 'https://cadence.sharepoint.com/sites/WWSPBAE-stg/Lists/Activity%20Report/Myitems.aspx';
  }
    
  

  public render(): React.ReactElement<IWwspbaeReactProps> {

    const { 
      nameValue, 
      nameSuggestions, 
      emailValue, 
      emailSuggestions,
      roleValue, 
      roleSuggestions 
    } = this.state;
    const nameInputProps = {
      placeholder: "Name",
      value: nameValue,
      onChange: this.onnameChange
    };
    const emailInputProps = {
      placeholder: "Email",
      value: emailValue,
      onChange: this.onEmailChange
    };
    const roleInputProps = {
      placeholder: "Role",
      value: roleValue,
      onChange: this.onRoleChange
    };

    return (
      <><><form id="container">
        
        <div className="container" >
          <header>Activity Report</header>
          <br />
          {/*<div className="stack-top">
          <ClipLoader css={override} size={150} color={'#123abc'} loading={this.state.loading} speedMultiplier={1} />
          </div>*/}
          <div className="row">
            <div className="col-sm-2 tasksInput">
              <label htmlFor="submitter">Submitter:</label>
            </div>
            <div className="col-sm-4">
              <input type="text" id="submitter" readOnly required className="form-control" />
            </div>
            <div className="col-sm-2 tasksInput">
              <label htmlFor="date">Date of Visit:</label>
            </div>
            <div className="col-sm-4">
              <input type="date" id="date" readOnly={this.state.hidden} required className="form-control" />
            </div>
          </div>
          <br />
          <div className="row">
            <div className="col-sm-2 tasksInput">
              <label htmlFor="dropdown_account">Account Name:</label>
            </div>
            <div className="col-sm-4">
            {/*<Searchable
                value="EFG/San Francisco" //if value is not item of options array, it would be ignored on mount
                placeholder="Select Account" // by default "Search"
                notFoundText="No result found" // by default "No result found"
                options={this.state.dropdown_items}
                onSelect={(option: any) => {
                  var array = option.label.split('/');
                  var acc = array[1];
                  console.log(array[1]);
                  $("#loc").val(acc); // as example - {value: '', label: 'All'}
                }}
                listMaxHeight={200} //by default 140
              />*/}
              {/*<Autocomplete
                id="dropdown"
                disabled = {this.state.hidden}
                options={this.state.dropdown_items_unique}
                getOptionLabel={(option) => option.label}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    variant="outlined"
                  />
                )}
                />*/}
              
              <Multiselect
                id="dropdown" 
                singleSelect
                options={this.state.dropdown_items_account} 
                isObject={false}
                onSelect={this.sort_location_based_on_selected_company}
                selectedValues={this.state.account_selected}
                disable={this.state.hidden}
                placeholder={'Select Account'}
              />
            </div>
            
            <div className="col-sm-2 tasksInput">
              <label htmlFor="loc" className="text-right">Site:</label>
            </div>
            <div className="col-sm-4">
              {/*<input type="text" id="loc" required readOnly className="form-control" />*/}
              {/*<Autocomplete
                id="location"
                disabled = {this.state.hidden}
                options={this.state.dropdown_items}
                getOptionLabel={(option) => option.label}
                multiple= {true.io
                //getOptionSelected={(option)=> option.label}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    variant="outlined"
                  />
                )}
                />*/}
                <Multiselect
                  id="location" 
                  options={this.state.dropdown_items_location_specific} 
                  isObject={false}
                  onSelect={this.fetchLocation}
                  selectedValues={this.state.readLocation}
                  disable={this.state.hidden}
                  ref={this.multiselectRef}
                  placeholder={'Select Site'}
                  />

            </div>
          </div>
          <br />
          <div className="row">
            <div className="col-sm-2 tasksInput">
              <label>Other Cadence Attendees:</label>
            </div>
            <div id="dv_custom" className="col-sm-10">
              <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={10}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={false}
                  required={true}
                  disabled={this.state.hidden}
                  onChange={this._getPeoplePickerItems}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} 
                  defaultSelectedUsers={this.state.defaultUser}              
                />
            </div>
          </div>
          <br />
          <div className="row">
            <div className="col-sm-2">
              <label>Purpose of the meeting:</label>
            </div>
            <div className="col-sm-5" id="myChecks">
              <input type="checkbox" disabled={this.state.hidden} name="location[]" id='lunch' defaultValue="Lunch and Learn" />
              <label htmlFor ='lunch'>Lunch and Learn</label>
              <br />
              <input type="checkbox" disabled={this.state.hidden} name="location[]" id='dis' defaultValue="Discovery" />
              <label htmlFor ='dis'>Discovery</label>
              <br />
              <input type="checkbox" disabled={this.state.hidden} name="location[]" id='sup' defaultValue="Support" />
              <label htmlFor ='sup'>Support</label>
              <br />
              <input type="checkbox" disabled={this.state.hidden} name="location[]" id='train' defaultValue="Training" />
              <label htmlFor ='train'>Training</label>
              <br /><br />
            </div>
          </div>
          <div className="row">
            <div className="col-sm-2 tasksInput">
              <label htmlFor="exc_sum">Executive Summary:</label>
            </div>
            <div className="col-sm-10">
              <textarea id="exc_sum" rows={4} required readOnly={this.state.hidden} className="form-control" defaultValue={""} />
            </div>
          </div>
          <br />
          <div className="row">
            <div className="col-sm-2 tasksInput">
              <label htmlFor="det_sum">Detailed Summary:</label>
            </div>
            <div className="col-sm-10">
              <textarea id="det_sum" rows={6} required readOnly={this.state.hidden} className="form-control" defaultValue={""} />
            </div>
          </div>
          <br />
          <div>
            <div className="row">
              <div className="col-sm-4 tasksInput">
                <label hidden={this.state.hidden}>Attendee's Details:</label>
              </div>
            </div>
            {/*
            <div className="row">
              <div className="col-sm-4 tasksInput">
                <input type="text" id="att_name" className="form-control" hidden={this.state.hidden} placeholder="Attendee's Name" />
              </div>
              <div className="col-sm-4 tasksInput">
                <input type="Email" id="att_email" className="form-control" hidden={this.state.hidden} placeholder="Attendee's Email" />
              </div>
              <div className="col-sm-4 tasksInput">
                <input type="text" id="att_title" className="form-control" hidden={this.state.hidden} placeholder="Attendee's Role" />
              </div>
            </div>
            <br />
              */}
            <div className="row">
              <div className="col-sm-4 tasksInput">
                <Autosuggest 
                    suggestions={nameSuggestions}
                    onSuggestionsFetchRequested={this.onnameSuggestionsFetchRequested}
                    onSuggestionsClearRequested={this.onnameSuggestionsClearRequested}
                    onSuggestionSelected={this.onnameSuggestionSelected}
                    getSuggestionValue={getSuggestionname}
                    renderSuggestion={renderSuggestion}
                    inputProps={nameInputProps}
                  />
              </div>
              <div className="col-sm-4 tasksInput">
                <Autosuggest 
                    suggestions={emailSuggestions}
                    onSuggestionsFetchRequested={this.onEmailSuggestionsFetchRequested}
                    onSuggestionsClearRequested={this.onEmailSuggestionsClearRequested}
                    onSuggestionSelected={this.onEmailSuggestionSelected}
                    getSuggestionValue={getSuggestionEmail}
                    renderSuggestion={renderSuggestion}
                    inputProps={emailInputProps}
                  />
              </div>
              <div className="col-sm-4 tasksInput">
                <Autosuggest 
                    suggestions={roleSuggestions}
                    onSuggestionsFetchRequested={this.onRoleSuggestionsFetchRequested}
                    onSuggestionsClearRequested={this.onRoleSuggestionsClearRequested}
                    onSuggestionSelected={this.onRoleSuggestionSelected}
                    getSuggestionValue={getSuggestionRole}
                    renderSuggestion={renderSuggestion}
                    inputProps={roleInputProps}
                  />
              </div>
            </div>
            <br />
            <div className="row">
              <div className="col-sm-4 text-center" />
              <div className="col-sm-4 text-center">
                <button type="button" id="add" className="btn btn-outline-info" hidden={this.state.hidden} onClick={this.addRow}>Add</button>
                &nbsp;
                <button type="button" id="delete" className="btn btn-outline-danger" hidden={this.state.hidden} onClick={this.deleteRow}>Remove</button>
              </div>
              <div className="col-sm-4 text-center" />
            </div>
            <br />
            <div>
              <table id="table" className="table border border-dark">
                <thead>
                  <tr className="table-dark text-dark">

                    <th />
                    <th scope="col">Attendee's Name</th>
                    <th scope="col">Attendee's Email</th>
                    <th scope="col">Attendee's Role</th>

                  </tr>
                </thead>
                <tbody id="tbody_saurabh">

                </tbody>
              </table>
            </div>
          </div>
          <br></br>
          <div className="row">
            <div className="col-sm-4 tasksInput">
              <label id = "notification_label" htmlFor="notification"></label>
            </div>
            <div className="col-sm-8">
              <p id="notification" className="notification form-control"></p>
            </div>
          </div>
          <br />

          <div className="row">
            <div className="col-sm-2 tasksInput">
              <label>Additional Recipients:</label>
            </div>
            <div id="notify" className="col-sm-10">
              <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={10}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={false}
                  required={false}
                  disabled={this.state.hidden}
                  onChange={this._getNotifyToPeople}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} 
                  defaultSelectedUsers={this.state.defaultNotifyTo}              
                />                
            </div>
            
          </div>
          
          <br />
        </div>
        <br></br>
        
        <div className="text-center">
          <button id="MyButton" type="button" className="btn btn-success" disabled={this.state.loading} hidden={this.state.hidden} onClick={() => this.createItem('Submitted')}>Submit</button>
          &nbsp;
          {/*<button type="button" className="btn btn-info" hidden={this.state.hidden} onClick={() => this.readItem()}>Read</button>
          &nbsp;*/}
          {/*<button type="button" className="btn btn-secondary" hidden={this.state.hidden} onClick={() => this.readOnlyItem()}>ReadOnly</button>
          &nbsp;*/}
          <button id="MyButton" type="button" className="btn btn-warning" disabled={this.state.loading} hidden={this.state.hidden} onClick={() => this.createItem('Saved')}>Save</button>
          &nbsp;
          <button id="MyButton" type="button" className="btn btn-danger" disabled={this.state.loading} hidden={this.state.hidden} onClick={this.closeForm}>Close</button>
          
        </div>
        {/*<br />
        <div className="text-center">
          <button type="button" className="btn btn-info" hidden={this.state.hidden} onClick={() => this.testItem()}>Test</button>
        </div>*/}
        <br></br>
        <div className='container'>
          <p id="validation" className="validation"></p>
          <div>
        
      </div>
        </div>

        
        
      </form>
      </>
      
      </>
  );
  
}  

}


