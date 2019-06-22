import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName, PlaceholderContent } from '@microsoft/sp-application-base';
//import { Dialog } from '@microsoft/sp-dialog';
//import * as strings from 'SpfxExtensionApplicationCustomizerStrings';

declare var $: any;
require('jquery');
require('modal');

import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  SPHttpClient,
  SPHttpClientResponse,
  //  SPHttpClientConfiguration,
  //  ISPHttpClientOptions
} from '@microsoft/sp-http';


export interface ISpfxExtensionApplicationCustomizerProperties {
  testMessage: string;
  userEmail: string;
  listName: string;
  itemId: number;
  url: string;
}

export interface IListItem {
  Title?: string;
  // EmailAddress: string;
  Id: number;
  //CompletedEnrollment: boolean;
  EnrollmentCompleted: Date;
}
const redirectUrl: string = 'https://6sc.sharepoint.com/sites/TPBC/SitePages/ThankYou.aspx';
const LOG_SOURCE: string = 'SpfxExtensionApplicationCustomizer';
export default class SpfxExtensionApplicationCustomizer extends BaseApplicationCustomizer<ISpfxExtensionApplicationCustomizerProperties> {
  private listItemEntityTypeName: string = undefined;
  private _topPlaceholder: PlaceholderContent | undefined;
  constructor() {
    super();
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css');
  }
  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Enrollments')/items?$orderby=Id desc&$top=1&$select=id&$filter=UserEmail+eq+'${this.properties.userEmail}'+and+EnrollmentCompleted+eq+null`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {
          return response.json();
        })
        .then((response: { value: { Id: number }[] }): void => {
          if (response.value.length === 0) {
            resolve(-1);
          }
          else {
            resolve(response.value[0].Id);
          }
        });
    });
  }

  private InsertJSFile(url: string): void {
    let head: any = document.getElementsByTagName("head")[0] || document.documentElement;
    let script: HTMLScriptElement = document.createElement("script");
    script.type = "text/javascript";
    script.src = url;
    head.appendChild(script);
    document.getElementsByTagName("head")[0].appendChild(script);
  }















  @override
  public onInit(): Promise<void> {

    Log.info(LOG_SOURCE, "Initialized ");
    this.properties.url = this.context.pageContext.web.absoluteUrl;
    this.properties.userEmail = this.context.pageContext.user.email.toString();
    this.properties.listName = "Enrollments";

    //this.InsertJSFile(`${this.context.pageContext.web.absoluteUrl}/SiteAssets/js/jquery-3.4.1.min.js`);
    //this.InsertJSFile(`https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js`);
    //this.InsertJSFile(`https://ajax.googleapis.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.js`);
    //this.InsertJSFile(`${this.properties.url}/SiteAssets/js/site.js`);
    this.InsertJSFile(`${this.context.pageContext.web.absoluteUrl}/SiteAssets/js/site.js`);

    //  only run for external user
    if (this.context.pageContext.user.isExternalGuestUser || this.context.pageContext.user.isAnonymousGuestUser) {
      // do not run on Thank you page or enrollment form
      if ((document.location.href).toLowerCase().indexOf("thankyou.aspx") == -1 && (document.location.href).toLowerCase().indexOf("enrollment.aspx") == -1) {
      //let restCall: string = `${this.properties.url}/_api/web/lists/getbytitle('Enrollments')/items?&$filter=UserEmail+eq+'${this.properties.userEmail}'+and+CompletedEnrollment+eq+1`;
      //this.ItemExists(restCall).then((result) => {
      this.getLatestItemId().then((result) => {
        this.properties.itemId = result;
        if (this.properties.itemId > 0) {    //  only update probably should add new item   ///   js/site.js
          //    this.InsertJSFile("https://6sc.sharepoint.com/sites/TPBC/SiteAssets/js/site.js");  // includes js in site assets library
          //this.showIframe(); // works for classic form or modern form or powerapp form   `https://web.powerapps.com/webplayer/iframeapp?hidenavbar=true&amp;screenColor=white&amp;appId=/providers/Microsoft.PowerApps/apps/473cc4ab-6455-463b-8f23-08a0ab89b856&amp;userEmail=${this.properties.userEmail}"`
          this.showForm();   /// displays form in modal
          this.InsertJSFile(`${this.properties.url}/SiteAssets/js/notEnrolled.js`);
        }
      })
        .catch((error: any) => {
          console.log(error);
          return true;  ///  log the error and return true so user can continue
        });
    }
  }
    return Promise.resolve();
  }

  public FormSave(): void {  
    let latestItemId: number = undefined;  
    this.updateStatus('Loading latest item...');  
    
    this.getLatestItemId()  
      .then((itemId: number): Promise<SPHttpClientResponse> => {  
        if (itemId === -1) {  
          throw new Error('No items found in the list');  
        }  
    
        latestItemId = itemId;  
        this.updateStatus(`Loading information about item ID: ${itemId}...`);  
          
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${latestItemId})?$select=Title,Id`,  
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'odata-version': ''  
            }  
          });  
      })  
      .then((response: SPHttpClientResponse): Promise<IListItem> => {  
        return response.json();  
      })  
      .then((item: IListItem): void => {  
        this.updateStatus(`Item ID1: ${item.Id}`);  
    
        const body: string = JSON.stringify({  
          'EnrollmentCompleted': `${new Date().toJSON()}`  
        });  
    
        this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${item.Id})`,  
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
          .then((response: SPHttpClientResponse): void => {  
            $.modal.defaults = { closeExisting: true };
            $.modal.close();
            this.updateStatus(`Item with ID: ${latestItemId} successfully updated`);  
          }, (error: any): void => {  
            this.updateStatus(`Error updating item: ${error}`);  
          });  
      });  
  }  
  private updateStatus(status: string, items: IListItem[] = []): void {  
    //this.domElement.querySelector('.status').innerHTML = status;  
    //this.updateItemsHtml(items);  
    console.log(status);
  }  










  
  public FormSave_old(): void {
    let etag: string = undefined;
    let listItemEntityTypeName: string = "SP.Data.EnrollmentsListItem";

    //const opt: ISPHttpClientOptions = { headers: { 'Content-Type': 'some value' }, body: { my: "bodyJson" } };
    //this.context.spHttpClient.post('', SPHttpClientConfiguration.);

    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${this.properties.itemId})?$select=Id`, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        etag = response.headers.get('ETag');
        return response.json();
      })
      .then((item: IListItem): Promise<SPHttpClientResponse> => {
        const postbody = JSON.stringify({__metadata:{'type':'SP.Data.EnrollmentsListItem'},Title:'test update'});
        //          'ContentTypeId': '0x0100DD259FFD1382FA4B8AD9BB7FE83F2C8A'
        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${item.Id})`, SPHttpClient.configurations.v1, {
          headers: {
            "accept": "application/json;odata=verbose",
            "content-Type": "application/json;odata=verbose",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
          },
          body: "{__metadata:{'type':'SP.Data.EnrollmentsListItem'},Title:'test update'}"
        });
      })    //,        'IF-MATCH': etag,        'X-HTTP-Method': 'MERGE'     data: "{__metadata:{'type':'SP.Data.EnrollmentsListItem'},Title:'thisisatest'}"              CompletedEnrollment: '1',         UserEmail: 'jd@projectpointinc.com'
      .then((response: SPHttpClientResponse): void => {
        $.modal.defaults = { closeExisting: true };
        $.modal.close();
      })
      .catch((error: any) => {
        console.log(error);
        //return true;  ///  log the error and return true so user can continue
      });

  }

  private showIframe(): void {
    let div: HTMLDivElement = document.createElement("div");
    div.id = "ex1";
    div.className = "modal";
    //div.onblur = ()=>{window.document.location.href=redirectUrl;};

    let iframe: any = document.createElement("iframe");
    iframe.onload = () => { this.iframeOnload(iframe.contentWindow.location.href); };
    //iframe.onloadend = () => {this.iframeOnload(iframe.contentWindow.location.href);};
    iframe.style.overflow = "hidden";
    iframe.width = "450";
    iframe.height = "650";
    iframe.frameBorder = "0";
    iframe.src = `${this.properties.url}/Lists/Enrollments/EditForm.aspx?isDlg=true&ID=${this.properties.itemId}&userEmail=${this.properties.userEmail}&Source=${this.properties.url}/Lists/Enrollments/EditForm.aspx?ID%3D${this.properties.itemId}&isDlg=true`;
    iframe.onblur = () => { window.location.href = redirectUrl; };
    div.appendChild(iframe);
    document.body.appendChild(div);

    $("#ex1").modal({
      escapeClose: false,
      clickClose: false,
      showClose: false,
      fadeDuration: 100
    });
  }


  public iframeOnload(url: string): void {
    if (url.indexOf("Source") == -1) {
      this.getLatestItemId().then((result) => {
        this.properties.itemId = result;
        if (this.properties.itemId > 0) {
          window.document.location.href = redirectUrl;
        } else {
          $(this).modal();
          $.modal.defaults = { closeExisting: true };
          $.modal.close();
        }
      })
        .catch((error: any) => {
          console.log(error);
          $.modal.defaults = { closeExisting: true };
          $.modal.close();  ///  log the error and user can continues
        });
    }
  }

  private showForm(): void {
    let formDiv: HTMLDivElement = document.createElement("div");
    formDiv.innerHTML = `
      <div id="ex1" class="modal">
        <input id="CompleteButton" type="button" value="Complete Enrollment" class="complete" />
        <input id="CancelButton" type="button" value="Cancel Enrollment" class="cancel" />
      </div>
      `;
    formDiv.querySelector('input.cancel').addEventListener('click', () => { this.FormCancel(); });  // error if not found
    formDiv.querySelector('input.complete').addEventListener('click', () => { this.FormSave(); });

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose: this._onDispose });
      if (this._topPlaceholder.domElement) {
        this._topPlaceholder.domElement.appendChild(formDiv);
      }
    }

    $("#ex1").modal({
      escapeClose: false,
      clickClose: false,
      showClose: false,
      fadeDuration: 100
    });
  }

  public FormCancel(): void {
    window.document.location.href = redirectUrl;
  }



  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
