import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'SpfxExtensionApplicationCustomizerStrings';

declare var $: any;
require('jquery');
require('modal');



import { SPComponentLoader } from '@microsoft/sp-loader';
/*  
bootstrap
  "externals": {
    "jquery": {
      "path": "https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js",
      "globalName": "jquery"
    },
    "bootstrap": {
      "path": "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js",
      "globalName": "bootstrap",
      "globalDependencies": ["jquery"]
    }
  },

  import('bootstrap');
*/


import {
  SPHttpClient,
  SPHttpClientResponse,
  SPHttpClientConfiguration,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

export interface ISpfxExtensionApplicationCustomizerProperties { testMessage: string;userEmail: string; listName:string; }

interface IListItem {
  Title?: string;
  EmailAddress:string;
  Id: number;
  CompletedEnrollment: boolean;
}

//declare function FormCancel(): void;

const LOG_SOURCE: string = 'SpfxExtensionApplicationCustomizer';
const redirectUrl: string = 'https://6sc.sharepoint.com/sites/TPBC/SitePages/ThankYou.aspx';

export default class SpfxExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxExtensionApplicationCustomizerProperties> {

    constructor() {
      super();
      SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css');
      //SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');
      //$('#ex1').on('hidden.bs.modal', function(e) {
      //  document.location.href = 'https://6sc.sharepoint.com/sites/TPBC/SitePages/ThankYou.aspx';
      //});
    }

   //public FormCancel() {window.location.href=redirectUrl;}

   private listItemEntityTypeName: string = undefined;

  private ItemExists(restCall: string): Promise<boolean> {
       return this.context.spHttpClient.get(restCall, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json: any) => {
        console.log(restCall);
        console.log(json.value.length);
        return json.value.length > 0;
      });
  }

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
  
  public CompleteEnrollment(): void {
    //this.updateStatus('Loading latest items...');
    $('#ex1').modal('hide');
    let latestItemId: number = undefined;
    let etag: string = undefined;
    let listItemEntityTypeName: string = undefined;
    this.getListItemEntityTypeName()   ///<d:ListItemEntityTypeFullName>SP.Data.EnrollmentsListItem</d:ListItemEntityTypeFullName>
      .then((listItemType: string): Promise<number> => {
        listItemEntityTypeName = listItemType;
        return this.getLatestItemId();
      })
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        //this.updateStatus(`Loading information about item ID: ${latestItemId}...`);
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${latestItemId})?$select=Id`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        etag = response.headers.get('ETag');
        return response.json();
      })
      .then((item: IListItem): Promise<SPHttpClientResponse> => {
        //this.updateStatus(`Updating item with ID: ${latestItemId}...`);
        const body: string = JSON.stringify({
          '__metadata': {
            'type': listItemEntityTypeName
          },
          'CompletedEnrollment': `${true}`
        });
        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${item.Id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': '',
              'IF-MATCH': etag,
              'X-HTTP-Method': 'MERGE'
            },
            body: body
          });
      })
      .then((response: SPHttpClientResponse): void => {
        //this.updateStatus(`Item with ID: ${latestItemId} successfully updated`);
      }, (error: any): void => {
        //this.updateStatus(`Error updating item: ${error}`);
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
    Log.info(LOG_SOURCE, "Initialized ${strings.Title}");

    let userEmail: string = this.context.pageContext.user.email.toString();   //  phase out
    this.properties.userEmail = this.context.pageContext.user.email.toString();
    this.properties.listName = "Enrollments";
    //  only run for external user
    if (this.context.pageContext.user.isExternalGuestUser || this.context.pageContext.user.isAnonymousGuestUser) {
      // do not run on Thank you page
      let url: string = this.context.pageContext.site.serverRequestPath.toString();
      if ((document.location.href).toLowerCase().indexOf("thankyou.aspx") == -1 && (document.location.href).toLowerCase().indexOf("enrollment.aspx") == -1)  {
        let restCall: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Enrollments')/items?&$filter=UserEmail+eq+'" + userEmail + "'+and+CompletedEnrollment+eq+1";

        //this.ItemExists(restCall).then((result) => {
        this.getLatestItemId().then((result) => {
          let itemId: number;
          itemId = result;
          // do not run if enrollment record found
          if (itemId > 0) {
            console.log("item exist .............     " + itemId);
            let message: string = this.properties.testMessage;
            if (!message) {
              message = '(No properties were provided.)';
            }

            let message2: string = "no placeholders";
            message2 = this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(", ");

            let modal: Element = document.createElement("div");
            //modal.innerHTML = "<div id='ex1' class='modal' style='z-index:1000;'><iframe style='overflow:hidden;' width='450' frameborder='0' height='650' src='https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/CompleteEnrollment.aspx?isDlg=true&ID="+itemId+"&userEmail="+userEmail+"'></iframe></div>";
            modal.innerHTML = "<div id='ex1' class='modal' style='z-index:1000;'><iframe style='overflow:hidden;' width='450' frameborder='0' height='650' src='https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?isDlg=true&ID="+itemId+"&userEmail="+userEmail+"'></iframe></div>";
            //modal.innerHTML = "<div id='ex1' class='modal' style='z-index:1000;'><iframe style='overflow:hidden;' width='450' frameborder='0' height='650' src='https://web.powerapps.com/webplayer/iframeapp?hidenavbar=true&amp;screenColor=white&amp;appId=/providers/Microsoft.PowerApps/apps/473cc4ab-6455-463b-8f23-08a0ab89b856&amp;userEmail="+userEmail+"'></iframe></div>";
            //modal.innerHTML += "<script type='text/javascript'>$('#ex1').blur(function() {window.location.href='"+redirectUrl+"'})</sript>";
            document.body.appendChild(modal);

            let head: any = document.getElementsByTagName("head")[0] || document.documentElement;
            let script: any = document.createElement("script");
            script.type = "text/javascript";
            script.appendChild(document.createTextNode("function FormCancel(){window.location.href='"+redirectUrl+"';}"));
            head.appendChild(script);
            script.appendChild(document.createTextNode("function FormSave(){window.location.href=window.location.href;}"));
            head.appendChild(script);
            script.appendChild(document.createTextNode("window.addEventListener('CloseDialog', function() { document.location.href='Home.aspx'; });"));
            head.appendChild(script);

            let jscript: Element = document.createElement("script");
            jscript.innerHTML = "window.addEventListener('CloseDialog', function() { document.location.href='Home.aspx'; });";
            document.head.appendChild(jscript);

            //script.appendChild(document.createTextNode("$('#ex1').blur(function() {window.location.href='"+redirectUrl+"'})"));
            //head.appendChild(script);

            //let modal: Element = document.createElement("div");
            //modal.innerHTML = "<div id='ex1' class='modal' style='z-index:1000;'><input type='button' value='Complete Enrollment' onclick='CompleteEnrollment();'/></div>";
            //document.body.appendChild(modal);

            window.addEventListener('CloseDialog', () => { window.location.href='Home.aspx'; });


            $('#ex1').modal('show');
            //$('#ex1').modal('hide');

            

            //$('#ex1').on('hide',() => {
            //  document.location.href=redirectUrl;
            //});
            //  works for bootstrap but not modal
            //$('#ex1').on('hidden.bs.modal',() => {
            //  this.FormCancel();
            //});

            // works for modal
            //$('#ex1').blur(function(){
            //  document.location.href=redirectUrl;
            //});

            //$('#ex1').focusout(function(){
            //  document.location.href=redirectUrl;
           //});

     


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

  public render(): void {

    window.addEventListener("CloseDialog", () => { document.location.href='Home.aspx'; });

}

}
