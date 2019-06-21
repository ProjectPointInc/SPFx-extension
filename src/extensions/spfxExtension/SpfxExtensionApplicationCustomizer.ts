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
  SPHttpClientConfiguration,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

export interface ISpfxExtensionApplicationCustomizerProperties {
  testMessage: string;
  userEmail: string;
  listName: string;
  itemId: number;
}

interface IListItem {
  // Title?: string;
  // EmailAddress: string;
  Id: number;
  CompletedEnrollment: boolean;
}

const LOG_SOURCE: string = 'SpfxExtensionApplicationCustomizer';
const redirectUrl: string = 'https://6sc.sharepoint.com/sites/TPBC/SitePages/ThankYou.aspx';

export default class SpfxExtensionApplicationCustomizer extends BaseApplicationCustomizer<ISpfxExtensionApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;




  constructor() {
    super();
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css');
  }

  public iframeOnload(url: string): void {
    if (url.indexOf("Source") == -1) {
      this.getLatestItemId().then((result) => {
        this.properties.itemId = result;
        if (this.properties.itemId > 0) {
          window.document.location.href = redirectUrl;
        } else {
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

  private listItemEntityTypeName: string = undefined;

  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Enrollments')/items?$orderby=Id desc&$top=1&$select=id&$filter=UserEmail+eq+'${this.properties.userEmail}'+and+CompletedEnrollment+eq+0`,
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
    iframe.src = `https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?isDlg=true&ID=${this.properties.itemId}&userEmail=${this.properties.userEmail}&Source=https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?ID%3D${this.properties.itemId}&isDlg=true`;
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

  private showForm(): void {
    let formDiv:HTMLDivElement = document.createElement("div");
    formDiv.innerHTML=`
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

  public FormSave(): void {
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
         const postbody = JSON.stringify({
          '__metadata': {'type': listItemEntityTypeName},
          'Title': 'test update',
          'CompletedEnrollment': '1',
          'UserEmail': 'jd@projectpointinc.com'
          });
              //          'ContentTypeId': '0x0100DD259FFD1382FA4B8AD9BB7FE83F2C8A'
        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${item.Id})`, SPHttpClient.configurations.v1, {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: postbody
        });
      })    //,        'IF-MATCH': etag,        'X-HTTP-Method': 'MERGE'
      .then((response: SPHttpClientResponse): void => {
        $.modal.defaults = { closeExisting: true };
        $.modal.close();
      })
      .catch((error: any) => {
        console.log(error);
        //return true;  ///  log the error and return true so user can continue
      });

  }




  private getListItemEntityTypeName(): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
      if (this.listItemEntityTypeName) {
        resolve(this.listItemEntityTypeName);
        return;
      }

      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')?$select=ListItemEntityTypeFullName`,
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
        });
    });
  }

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized ");
    let userEmail: string = this.context.pageContext.user.email.toString();   //  phase out
    this.properties.userEmail = this.context.pageContext.user.email.toString();
    this.properties.listName = "Enrollments";
    //  only run for external user
    if (this.context.pageContext.user.isExternalGuestUser || this.context.pageContext.user.isAnonymousGuestUser) {
      // do not run on Thank you page
      let url: string = this.context.pageContext.site.serverRequestPath.toString();
      if ((document.location.href).toLowerCase().indexOf("thankyou.aspx") == -1 && (document.location.href).toLowerCase().indexOf("enrollment.aspx") == -1) {
        let restCall: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Enrollments')/items?&$filter=UserEmail+eq+'" + userEmail + "'+and+CompletedEnrollment+eq+1";
        //this.ItemExists(restCall).then((result) => {
        this.getLatestItemId().then((result) => {
          this.properties.itemId = result;
          if (this.properties.itemId > 0) {    //  only update probably should add new item
            //this.InsertJSFile("https://6sc.sharepoint.com/sites/TPBC/SiteAssets/SpfxExtensionApplicationCustomizerCustom.js");  // includes js in site assets library
            this.showIframe(); // works for classic form or modern form or powerapp form   `https://web.powerapps.com/webplayer/iframeapp?hidenavbar=true&amp;screenColor=white&amp;appId=/providers/Microsoft.PowerApps/apps/473cc4ab-6455-463b-8f23-08a0ab89b856&amp;userEmail=${this.properties.userEmail}"`
            //this.showForm();   /// displays form in modal
          }
        })
          .catch((error: any) => {
            console.log(error);
            return true;  ///  log the error and return true so user can continue
          });
      //}
    //}
    return Promise.resolve();
  }

  public render(): void {
  }
  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
