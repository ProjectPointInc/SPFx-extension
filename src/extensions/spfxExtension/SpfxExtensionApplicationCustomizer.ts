import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'SpfxExtensionApplicationCustomizerStrings';
import * as $ from 'jquery';
import {
  SPHttpClient,
  SPHttpClientResponse,
  SPHttpClientConfiguration,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

export interface ISpfxExtensionApplicationCustomizerProperties { testMessage: string; }

const LOG_SOURCE: string = 'SpfxExtensionApplicationCustomizer';

export default class SpfxExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxExtensionApplicationCustomizerProperties> {

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

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized ${strings.Title}");

    // do not run on Thank you page
    let url: string = this.context.pageContext.site.serverRequestPath.toString();
    //if (url.search(/ThankYou.aspx/gi) == -1) {    ///  does not work in some cases (some environments?)   added line below too
      if ((document.location.href).toLowerCase().indexOf("thankyou.aspx") == -1) {
        let userEmail: string = this.context.pageContext.user.email.toString();
        // debugging
        //let restCall: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('Enrollments')/items";
        //let restCall: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Enrollments')/items?&$filter=UserEmail+eq+'test'";
        let restCall: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Enrollments')/items?&$filter=UserEmail+eq+'" + userEmail + "'";

        this.ItemExists(restCall).then((result) => {
          let itemExists: boolean;
          itemExists = result;
          // do not run if enrollment record found
          if (!itemExists) {
            console.log("item exist .............     " + itemExists);
            let message: string = this.properties.testMessage;
            if (!message) {
              message = '(No properties were provided.)';
            }

            let message2: string = "no placeholders";
            message2 = this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(", ");

            // debugging
            //Dialog.alert(`Title:${strings.Title}    QueryParam:${message}    Available Place Holders:${message2}`);
            //alert(`userEmail: ${userEmail}    QueryParam:${message}    Available Place Holders:${message2}     ListItem Exists:${itemExists}`);

            let appendHTML: Element = document.createElement("div");
            let formGUID: string = "fca15810-833c-48a2-b45b-14d4a16382e3";     //"473cc4ab-6455-463b-8f23-08a0ab89b856";   /// need to add to environment variables
            let formURL: string = "https://web.powerapps.com/webplayer/iframeapp?hidenavbar=false&amp;screenColor=white&amp;appId=/providers/Microsoft.PowerApps/apps/" + formGUID + "&amp;userEmail=" + userEmail;
            appendHTML.innerHTML = '<div style="position:absolute;width:444px;height:790px;z-index:15;top:25%;left:50%;margin:-200px 0 0 -200px;border:1px solid black;"><iframe width="444" height="790" src="' + formURL + '"></iframe></div>';
            document.body.appendChild(appendHTML);
          }
        })
          .catch((error: any) => {
            console.log(error);
            return true;  ///  log the error and return true so user can continue
          });
      }
    //}
    return Promise.resolve();
  }
}

/*
 this is not a good solution:  better to add js to project assets and deploy to CDN

 however great to inject html elements with HTMLElement

     let articleRedirectScriptTag: HTMLScriptElement = document.createElement("script");
       articleRedirectScriptTag.src = "https://*******.sharepoint.com/sites/CommunicationSiteTopic/Shared%20Documents/MyScript.js";
       articleRedirectScriptTag.type = "text/javascript";
       document.body.appendChild(articleRedirectScriptTag);


       https://projectpoint.sharepoint.com/sites/dev/_layouts/15/listform.aspx?PageType=8&ListId=%7BDA1B998E-9B6B-446B-84D8-A5A375E1A088%7D&RootFolder=%2Fsites%2Fdev%2FLists%2FEnrollments&Source=https%3A%2F%2Fprojectpoint.sharepoint.com%2Fsites%2Fdev%2FLists%2FEnrollments%2FAllItems.aspx&ContentTypeId=0x0100DD259FFD1382FA4B8AD9BB7FE83F2C8A

<script>SP.UI.ModalDialog.showModalDialog('https://projectpoint.sharepoint.com/sites/dev/Lists/Enrollments/NewForm.aspx?Source=https%3A%2F%2Fprojectpoint%2Esharepoint%2Ecom%2Fsites%2Fdev%2FSitePages%2FThankYou%2Easpx');</script>

<iframe width="100%" height="100%" src="https://web.powerapps.com/webplayer/iframeapp?hidenavbar=true&amp;screenColor=white&amp;appId=/providers/Microsoft.PowerApps/apps/473cc4ab-6455-463b-8f23-08a0ab89b856"></iframe>


6sc form guid
fca15810-833c-48a2-b45b-14d4a16382e3

pp form guid
473cc4ab-6455-463b-8f23-08a0ab89b856

*/
