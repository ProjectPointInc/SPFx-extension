import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'SpfxExtensionApplicationCustomizerStrings';

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

  constructor() {
    super();
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jquery-modal/0.9.1/jquery.modal.min.css');
  }

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized ${strings.Title}");
    let userEmail: string = this.context.pageContext.user.email.toString();
    //  only run for external user
    if (this.context.pageContext.user.isExternalGuestUser || this.context.pageContext.user.isAnonymousGuestUser) {
      // do not run on Thank you page
      let url: string = this.context.pageContext.site.serverRequestPath.toString();
      if ((document.location.href).toLowerCase().indexOf("thankyou.aspx") == -1) {    
        let restCall: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Enrollments')/items?&$filter=UserEmail+eq+'" + userEmail + "'+and+CompletedEnrollment+eq+1";

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

           let modal: Element = document.createElement("div");
           modal.innerHTML = `
             <div id="ex1" class="modal">
               <iframe width="450"frameborder="0" height="650" src="https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/completeEnrollment.aspx"></iframe>
             </div>`;

           document.body.appendChild(modal);
           $('#ex1').modal('show');

           $('#ex1').focusout(function() {
              document.location.href='https://6sc.sharepoint.com/sites/TPBC/SitePages/ThankYou.aspx';
           });
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
}
