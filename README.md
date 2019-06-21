## spfx-extension

SPFx project (extension feature)

Inserts a jQuery popup on SPO and Modern pages - replaced with iframe containing powerapp form (to do replace iframe with overlay host

Scoped to the site collection

Assets are deployed to CDN

## TO DO

replace iframe popup with KnockOut overlay 

powerapp form navigation



##  Deployment steps

Deploy the App to SPO
 - Copy spfx-extension.sppkg to App Catalog (SPO->AppCatalog site collection->Apps for SharePoint library)
 - Deploy and allow permissions for the App (in Apps for SharePoint library)
 
Add the App to Site Collection
 - In the target site collection, go to site contents then "add 
an app"
 - Select spfx-extension-client-side-solution
 
Done


### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO

###  Bundle

gulp bundle --ship

gulp package-solution --ship


###  Check In Original

git commit -a -m "message"

git push origin master

### New Branch

git checkout -b modalform

git push -u origin modalform


###  CDN Enable on tenant

Get-SPOTenantCdnEnabled -CdnType Public

Get-SPOTenantCdnOrigins -CdnType Public

Get-SPOTenantCdnPolicies -CdnType Public

Set-SPOTenantCdnEnabled -CdnType Public




###  JUNK



    //let latestItemId: number = undefined;
    //let etag: string = undefined;
    //let postbody: string = undefined;
    //let listItemEntityTypeName: string = "SP.Data.EnrollmentsListItem";
    /*  this.getListItemEntityTypeName()   ///<d:ListItemEntityTypeFullName>SP.Data.EnrollmentsListItem</d:ListItemEntityTypeFullName>
        .then((listItemType: string): Promise<number> => {
          listItemEntityTypeName = listItemType;
          return this.getLatestItemId();
        })
        .then((itemId: number): Promise<SPHttpClientResponse> => {
          if (itemId === -1) {
            throw new Error('No items found in the list');
          }
          latestItemId = itemId;
  
          return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${latestItemId})?$select=Id`,
      */
       //          '__metadata': {
  //  'type': 'SP.Data.EnrollmentsListItem'
  //},
  //'IF-MATCH': etag,

/*

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

interface IListItem {
  Title?: string;
  EmailAddress: string;
  Id: number;
  CompletedEnrollment: boolean;
}
*/
    //SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');
    //$('#ex1').on('hidden.bs.modal', function(e) {
    //  document.location.href = 'https://6sc.sharepoint.com/sites/TPBC/SitePages/ThankYou.aspx';
    //});

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
*/

  /*
  private _renderPlaceHolders(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
      if (this._topPlaceholder.domElement) {
        //this._topPlaceholder.domElement.innerHTML = `<div id="ex1" class="modal"><iframe id="mfWindow"  onLoad="this.iframeOnload()" style="overflow:hidden;" width="450" frameborder="0" height="650" src="https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?isDlg=true&ID=${itemId}&userEmail=${userEmail}&Source=https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?closemeplease&ID=${itemId}"></iframe></div>`;
       // this._topPlaceholder.domElement.innerHTML = `<div id="ex1" class="modal"><iframe id="mfWindow" onLoad="${this.iframeOnload()}" style="overflow:hidden;" width="450" frameborder="0" height="650" src="https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?isDlg=true&ID=71&userEmail=jon92651@gmail.com&Source=https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?closemeplease&ID=71"></iframe></div>`;
        this._topPlaceholder.domElement.innerHTML = `<div id="ex1" class="modal"><iframe id="mfWindow" onLoad="iframeOnload2(this.contentWindow.location)" style="overflow:hidden;" width="450" frameborder="0" height="650" src="https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?isDlg=true&ID=71&userEmail=jon92651@gmail.com&Source=https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?closemeplease&ID=71"></iframe></div>`;
      }
    }

  }
 
  
            //this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
            //  on dom load
            //$(()=>{
               // iframeOnload2('2');
            //});


            

            //$('#ex1').modal('show');
             
let test: GlobalEventHandlers;
test.addEventListener() 
            let iframeEvent: EventListenerOrEventListenerObject  = window.addEventListener('CloseDialog', () => window.location.href='Home.aspx' );

            //mfWindow.addEventListener('CloseDialog', (e:Event) => window.location.href='Home.aspx' );
      
            $(()=>{
              this.iframeOnload("test");
            });

            let modal: Element = document.createElement("div");
            //modal.innerHTML = "<div id='ex1' class='modal' style='z-index:1000;'><iframe style='overflow:hidden;' width='450' frameborder='0' height='650' src='https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/CompleteEnrollment.aspx?isDlg=true&ID="+itemId+"&userEmail="+userEmail+"'></iframe></div>";
            modal.innerHTML = `<div id="ex1" class="modal"><iframe id="mfWindow"  onLoad="this.iframeOnload()" style="overflow:hidden;" width="450" frameborder="0" height="650" src="https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?isDlg=true&ID=${itemId}&userEmail=${userEmail}&Source=https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?closemeplease&ID=${itemId}"></iframe></div>`;
            //modal.innerHTML = "<div id='ex1' class='modal' style='z-index:1000;'><iframe style='overflow:hidden;' width='450' frameborder='0' height='650' src='https://web.powerapps.com/webplayer/iframeapp?hidenavbar=true&amp;screenColor=white&amp;appId=/providers/Microsoft.PowerApps/apps/473cc4ab-6455-463b-8f23-08a0ab89b856&amp;userEmail="+userEmail+"'></iframe></div>";
            //modal.innerHTML += "<script type='text/javascript'>$('#ex1').blur(function() {window.location.href='"+redirectUrl+"'})</sript>";
            document.body.appendChild(modal);




            //script.appendChild(document.createTextNode("function FormSave(){window.location.href=window.location.href;}"));
            //head.appendChild(script);
            //script.appendChild(document.createTextNode(`function iframeOnload(url){alert(url.indexOf('Source'));}`));
            //head.appendChild(script);



     

                       
            console.log("item exist .............     " + itemId);
            let message: string = this.properties.testMessage;
            if (!message) {
              message = '(No properties were provided.)';
            }

            let message2: string = "no placeholders";
            message2 = this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(", ");
  
                   function iframeOnload(){alert('2');}
                   alert(iframeOnload());
     
     //"function iframeOnload(url){alert(url);alert(url.indexOf('Source');if (url.indexOf('Source') == -1){alert(); $('#ex1').modal('hide');}"
     
                   $('#mfWindow').onLoad( () => {
                   alert($('#mfWindow').contentWindow.location);
                   alert($('#mfWindow').contentWindow.location.indexOf('Source'));
                   //if (url.indexOf('closemeplease') != -1){ $('#ex1').modal('hide');  }
                 });
     
                 });
     
   












            //javascript:if (this.contentWindow.location.indexOf("closemeplease") != -1){ alert()  }
            //if (url.indexOf('closemeplease') != -1){$('#ex1').modal('hide');}
            //window.addEventListener('CloseDialog', function() { document.location.href='Home.aspx'; });
            //let jscript: Element = document.createElement("script");
            //jscript.innerHTML = "window.addEventListener('CloseDialog', function() { document.location.href='Home.aspx'; });";
            //document.head.appendChild(jscript);

            //script.appendChild(document.createTextNode("$('#ex1').blur(function() {window.location.href='"+redirectUrl+"'})"));
            //head.appendChild(script);

            //let modal: Element = document.createElement("div");
            //modal.innerHTML = "<div id='ex1' class='modal' style='z-index:1000;'><input type='button' value='Complete Enrollment' onclick='CompleteEnrollment();'/></div>";
            //document.body.appendChild(modal);

            //window.addEventListener('CloseDialog', () => window.location.href='Home.aspx' );

            //mfWindow.addEventListener('CloseDialog', (e:Event) => window.location.href='Home.aspx' );
            //$('#ex1').modal('hide');
            //$('#ex1').on('hide',() => {
            //  document.location.href=redirectUrl;
            //});
            //  works for bootstrap but not modal
            //$('#ex1').on('hidden.bs.modal',() => {
            //  this.FormCancel();
            //});

            // works for modal
            //$('#ex1').modal.blur(()=>{
            //  window.location.href=redirectUrl;
            //});

            //$('#ex1').focusout(function(){
            //  document.location.href=redirectUrl;
            //});

  
    public render(): void {
  
      //window.addEventListener('CloseDialog', (e:Event) => window.location.href='Home.aspx' );
  
      let itemId: number = 71;
      let userEmail: string = "test";
  
      let modal: Element = document.createElement("div");
      //modal.innerHTML = "<div id='ex1' class='modal' style='z-index:1000;'><iframe style='overflow:hidden;' width='450' frameborder='0' height='650' src='https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/CompleteEnrollment.aspx?isDlg=true&ID="+itemId+"&userEmail="+userEmail+"'></iframe></div>";
      modal.innerHTML = `<div id="ex1" class="modal"><iframe id="mfWindow"  onLoad="{()=>this.iframeOnload(this.contentWindow.location)}" style="overflow:hidden;" width="450" frameborder="0" height="650" src="https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?isDlg=true&ID=${itemId}&userEmail=${userEmail}&Source=https://6sc.sharepoint.com/sites/TPBC/Lists/Enrollments/EditEnrollment.aspx?closemeplease&ID=${itemId}"></iframe></div>`;
      //modal.innerHTML = "<div id='ex1' class='modal' style='z-index:1000;'><iframe style='overflow:hidden;' width='450' frameborder='0' height='650' src='https://web.powerapps.com/webplayer/iframeapp?hidenavbar=true&amp;screenColor=white&amp;appId=/providers/Microsoft.PowerApps/apps/473cc4ab-6455-463b-8f23-08a0ab89b856&amp;userEmail="+userEmail+"'></iframe></div>";
      //modal.innerHTML += "<script type='text/javascript'>$('#ex1').blur(function() {window.location.href='"+redirectUrl+"'})</sript>";
      document.body.appendChild(modal);
  
      $('#ex1').modal('show');
  
  }
  */

