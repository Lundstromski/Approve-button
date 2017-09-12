import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  CellFormatter,
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';
import { SPPermission } from "@microsoft/sp-page-context";
import pnp, { List, ItemUpdateResult, Item } from 'sp-pnp-js';
//import pnp from "sp-pnp-js";

import * as strings from 'toggleStrings';
import Toggle from './components/Toggle';
import { IToggleProps } from './components/IToggleProps';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IToggleProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'ToggleFieldCustomizer';

export default class ToggleFieldCustomizer
  extends BaseFieldCustomizer<IToggleProperties> {

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated ToggleFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "Toggle" and "${strings.Title}"`);
    return Promise.resolve<void>();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.  The CellFormatter is a utility
    // that you can use to convert the cellValue to a text string.
    const value: string = event.cellValue;
    const id: string = event.row.getValueByName('ID').toString();
    const hasPermissions: boolean = this.context.pageContext.list.permissions.hasPermission(SPPermission.editListItems);

    const toggle: React.ReactElement<{}> =
      React.createElement(Toggle, { checked: value, id: id, disabled: !hasPermissions, onChanged: this.onToggleValueChanged.bind(this) } as IToggleProps);

    ReactDOM.render(toggle, event.cellDiv);

  }

  // function getCurrentUser() {
  //   var context = new SP.ClientContext.get_current();
  //   var web = context.get_web();
  //   currentUser = web.get_currentUser();
  //   context.load(currentUser);
  //   context.executeQueryAsync(onSuccessMethod, onRequestFail);
  // }
  // function onSuccessMethod(sender, args) {
  //   var account = currentUser.get_loginName();
  //   var currentUserAccount = account.substring(account.indexOf("|") + 1);
  //   alert(currentUserAccount);
  // }
  // // This function runs if the executeQueryAsync call fails.
  // function onRequestFail(sender, args) {
  //   alert('request failed' + args.get_message() + '\n' + args.get_stackTrace());
  // }

  // @override
  // public getCurrentUser() {
  //   pnp.sp.web.currentUser.get()
  // }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.cellDiv);
    super.onDisposeCell(event);
  }
  
  private onToggleValueChanged(value: boolean, id: string): void {

    var utc = new Date().toJSON().slice(0,10).replace(/-/g,'/');

    // var currentDate = new Date()
    // var day = currentDate.getDate()
    // var month = currentDate.getMonth() + 1
    // var year = currentDate.getFullYear()
    // var time = day + "/" + month + "/" + year;

    // function getCurrentUser() {
    //   var context = new SP.ClientContext.get_current();
    //   var web = context.get_web();
    //   currentUser = web.get_currentUser();
    //   context.load(currentUser);
    //   context.executeQueryAsync(onSuccessMethod, onRequestFail);
    // }
    // function onSuccessMethod(sender, args) {
    //   var account = currentUser.get_loginName();
    //   var currentUserAccount = account.substring(account.indexOf("|") + 1);
    //   alert(currentUserAccount);
    // }
    // // This function runs if the executeQueryAsync call fails.
    // function onRequestFail(sender, args) {
    //   alert('request failed' + args.get_message() + '\n' + args.get_stackTrace());
    // }
    //var currentUser = pnp.sp.web.currentUser;
    var account = this.context.pageContext.user.loginName;
    var currentUserAccount = account.substring(account.indexOf('|') + 1);
    console.log((currentUserAccount));



    var User = this.context.pageContext.user.displayName;
    //console.log(User);

    let etag: string = undefined;
    pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).get(undefined, {
      headers: {
        'Accept': 'application/json;odata=minimalmetadata'
      }
    })
      .then((item: Item): Promise<any> => {
        etag = item["odata.etag"];
        return Promise.resolve((item as any) as any);
      })
      .then((item: any): Promise<ItemUpdateResult> => {
        let updateObj: any = {};
        updateObj[this.context.field.internalName] = value;
        return pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).update(updateObj, etag);
      })
      .then((result: ItemUpdateResult): void => {
        
        //if (pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).fields.getByTitle('Test_Single_Line_Of_Text').get.toString() == User) {  //if the item already has a approved-value
          //console.log('==User');
          pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).update({
            ApprovedAt: utc //Displays the date correctly
            }).then(r => {
          });

          pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).update({
            //ApprovedBy: User //this.context.pageContext.user.displayName //pnp.sp.web.currentUser
            Test_Single_Line_Of_Text: currentUserAccount //Displays User correctly
            }).then(r => {
          });

          pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).update({
             ApprovedBy: currentUserAccount //Do not display User correctly
          }).then(r => {
          });

        // } else {
        //   pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).update({
        //     //pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).update(pnp.sp.web.fields.getByTitle('ApprovedBy'), this.context.pageContext.user.displayName);
        //     ApprovedAt: utc //Displays the date correctly
        //     }).then(r => {  
        //     });
        //     console.log('!=User');
        //     pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).update({
        //     //ApprovedBy: User //this.context.pageContext.user.displayName //pnp.sp.web.currentUser
        //     Test_Single_Line_Of_Text: User //Displays User correctly
        //     }).then(r => {                
        //     });

        //     pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).update({
        //       //Test_Single_Line_Of_Text: User //Displays User correctly
        //       ApprovedBy: User
        //     }).then(r => {
        //     });

        // }//if 

        //console.log(this.context.pageContext.user.displayName); //Displays username correctly
        //console.log('Time: ' + time);
        //console.log(`Item with ID: ${id} successfully updated`);
      }, (error: any): void => {
        console.log('Loading latest item failed with error: ' + error);
      });
  }
}
