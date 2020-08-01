import * as React from 'react';
import {SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';
import { DatePicker, DayOfWeek, IDatePickerStrings, mergeStyleSets, PrimaryButton } from '@fluentui/react';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { CompactPeoplePicker, IBasePickerSuggestionsProps, ValidationState } from '@fluentui/react/lib/Pickers';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { assign } from '@microsoft/sp-lodash-subset';
import {IListItem} from './IListItem';
import {TextField} from '@fluentui/react/lib/TextField';
export interface INewFormProps
{
  siteUrl:string;
  spHttpClient:SPHttpClient;
  maxNrOfUsers?:number;
}


const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 300 },
  };
  const suggestionProps: IBasePickerSuggestionsProps={
    suggestionsHeaderText: 'Sugessted People'
  };
  const limitedSearchAdditionalProps:IBasePickerSuggestionsProps={
   
   
    searchForMoreText: 'Search For More'
  };
  
const options: IDropdownOption[] =[
    {key :'Conference', text:'Conference'},
    {key :'Training', text:'Training'},
    {key :'Town Hall Meeting', text:'Town Hall Meeting'},
  ];
  

  export interface INewFormState{
    Title?:string;
    EventDate?:Date;
    Organizer?:string;
    EventDetails?:string;
    EventType?:string;
    UserID?:string;
  }
  const limitedSearchSuggestionProps:IBasePickerSuggestionsProps = assign(limitedSearchAdditionalProps,suggestionProps);
  export class NewForm extends React.Component<INewFormProps,INewFormState,{}>{
    private listItemEntityTypeName: string = undefined;
    constructor(props:INewFormProps,state:INewFormState)
    {
      super(props);
      this.state={
        Title:"",
        EventDate:null,
        Organizer:"",
        EventDetails:"",
        EventType:"",
        UserID:'0'
      };
    }
    public render():React.ReactElement<INewFormProps>
    {
      return(
          <div>
        <div>Welcome to New Form</div>
        <label>Event Name</label>
        <TextField
            id="txt_eventname"
            title={this.state.Title}
            placeholder="Please enter event name..."
            onChange={(event,value)=>{this.setState({Title:value});}}
            />
            <label>Event Details</label>
          <TextField
            id="txt_eventdetails"
            title={this.state.EventDetails}
            placeholder="Please enter event details..."
            multiline rows={5}
            onChange={(event,value)=>{this.setState({EventDetails:value});}}
            />
            <label>Event Date</label>
            <DatePicker
            id="dt_eventdate"
            placeholder="Select a date"
            onSelectDate={date=> this.setState({EventDate:date})}
            />
            <label>Organizer</label>
            <CompactPeoplePicker
            onResolveSuggestions={this.onResolveSuggestions}
            pickerSuggestionsProps={limitedSearchSuggestionProps}
            className={'ms-PeoplePicker'}
            itemLimit={1}
            inputProps={{
              onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
              onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
              'aria-label': 'People Picker',
            }}
            resolveDelay={300}
            onChange={(this._onChangePeoplePicker.bind(this))}
          />
    
            <label>Event Type</label>
            <Dropdown id="drd_eventtype"
            placeholder="Select the Event Type"
            options={options}
            styles={dropdownStyles}
            onChange={(event,value)=>{this.setState({EventType:value.text});}}
            />
            <label></label>
            <PrimaryButton text="Save" onClick={this._SaveNewItem.bind(this)}/>
    
          </div>
        );
      }
      private _onChangePeoplePicker = (items?:any) =>{
        let loginname ="";
        let UserID:string = "";
        items.map(item=>(
          loginname = item.itemID
        ));
        this._getUserID(loginname);
      }
      private _getUserID(userAccountName):void{
        if(userAccountName)
        {
          userAccountName = "'" + userAccountName + "'";
          userAccountName = encodeURIComponent(userAccountName);
          const url:string = `${this.props.siteUrl}/_api/web/siteusers(@v)?@v=${userAccountName}`;
            console.log(url);
            this.props.spHttpClient.get(url,SPHttpClient.configurations.v1,
                {
                  headers: {
                    'Accept': 'application/json;odata=verbose',
                    'odata-version': ''
                  }
                })
            .then((response:SPHttpClientResponse)=>{
              return response.json();
            },(error:any):void=>{
              console.log('error');
            })
            .then((jsonresponse:any)=>{
              let userid:string =  jsonresponse.d.Id;
              this.setState(
                {Organizer:userid}
                );
            });
      
        }
      }
      private _SaveNewItem(event):void{
        let Title = this.state.Title;
        let EventDetails = this.state.EventDetails;
        let EventDate = this.state.EventDate;
        let EventType = this.state.EventType;
      
        let Organizer = this.state.Organizer;
      
        this.getListItemEntityTypeName()
            .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
              const body: string = JSON.stringify({
                '__metadata':
                {
                  'type': listItemEntityTypeName
                },
                'Title': Title,
                'EventDetails':EventDetails,
                'OrganizerId':Number(Organizer), //{'results':[Organizer]},
                'EventDate': EventDate.toISOString(),
                //'EventType': {"__metadata":{"type":"Collection(Edm.String)"},"results":[EventType]}  // multi selector
                'EventType': EventType
              });
              return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('SPFxEvents')/items`,
                SPHttpClient.configurations.v1,
                {
                  headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-type': 'application/json;odata=verbose',
                    'odata-version': ''
                  },
                  body: body
                });
    
            })
            .then((response: SPHttpClientResponse): Promise<IListItem> => {
              return response.json();
              alert ('Iteme sucussefull');
              window.location.reload(true);
            });
        }
        
        private onResolveSuggestions = (searchText: string, currentPersonas: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> => {
            return this.onFilterChanged(searchText, currentPersonas, this.props.maxNrOfUsers);
          }
          private onFilterChanged = (filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {
            return new Promise<IPersonaProps[]>((resolve,reject) => {
                let filteredPersonas: IPersonaProps[] = [];
         
            this.props.spHttpClient.get(`${this.props.siteUrl}/_api/search/query?querytext='*${filterText}*'&rowlimit=10&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'odata-version': ''
              }
            })
            .then((response:SPHttpClientResponse):Promise<{PrimaryQueryResult:any}> =>{
              return response.json();
            })
            .then((response:{PrimaryQueryResult:any}):void =>{
                let revelantResults:any = response.PrimaryQueryResult.RelevantResults;
                let resultCount:number = revelantResults.TotalRows;
                if(resultCount>0)
                {
                  revelantResults.Table.Rows.forEach( (row)=>{
                    let persona:IPersonaProps ={};
                    row.Cells.forEach((cell)=>{
                      if(cell.Key === 'JobTitle')
                      persona.secondaryText = cell.Value;
                      if(cell.Key === 'PictureURL')
                      persona.imageUrl = cell.Value;
                      if(cell.Key === 'PreferredName')
                      persona.primaryText = cell.Value;
                      if(cell.Key === 'AccountName')
                      persona.itemID = cell.Value;
                    });
                    filteredPersonas.push(persona);
                  });
                }
               resolve (filteredPersonas);
               
              },
              
              (error:any):void =>{
                reject();
            });
         
          });
          }

          private getListItemEntityTypeName(): Promise<string> {
            return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
              if (this.listItemEntityTypeName) {
                resolve(this.listItemEntityTypeName);
                return;
              }
        
              this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('SPFxEvents')?$select=ListItemEntityTypeFullName`,
                SPHttpClient.configurations.v1,
                {
                  headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                  }
                })
                .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
                  return response.json();
                }, (error: any): void => {
                  reject(error);
                })
                .then((response: { ListItemEntityTypeFullName: string }): void => {
                  this.listItemEntityTypeName = response.ListItemEntityTypeFullName;
                  resolve(this.listItemEntityTypeName);
                  alert ('Iteme sucussefull');
                });
            });
          }
        }
        
        


    
    
 
  
  
    
  
  

  


  


  
 
  
  
  
  


   
    

  

