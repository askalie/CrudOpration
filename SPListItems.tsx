import * as React from 'react';
import {IListItem} from './IListItem';
import { List } from '@fluentui/react/lib/List';

export interface IListItemsProps{
  spItems:IListItem[];
}

export class SPListItems extends React.Component<IListItemsProps,{}>{
  public render():React.ReactElement<IListItemsProps>{
    return(
      <div>
        <List items={this.props.spItems} onRenderCell={this.onRenderCell}/>
        <ul>
          {
            this.props.spItems.map(spItem =>(
              <li>
                ID: {spItem.Id} - Descrption : {spItem.Title} - Date : {spItem.EventDate} - Modified: {spItem.EventType}
              </li>
            ))
           } 
        </ul>
        </div>
    );
  }

  private onRenderCell = (items:IListItem, index:number |undefined):JSX.Element=>
  {
    return(<div>
      {items.Id} {items.Title} {items.EventDetails} {items.EventDate} {items.Organizer} {items.EventType}
    </div>);
  }
}
