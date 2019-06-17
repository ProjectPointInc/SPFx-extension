## spfx-extension

SPFx project (extension feature)

Inserts a jQuery popup on SPO and Modern pages

Scoped to the site collection

Assets are deployed to CDN

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


###  Check In

git commit -a -m "message"

git push origin master

### New Branch

git checkout -b feature_branch_name

git push -u origin feature_branch_name


###  CDN Enable on tenant

Get-SPOTenantCdnEnabled -CdnType Public

Get-SPOTenantCdnOrigins -CdnType Public

Get-SPOTenantCdnPolicies -CdnType Public


Set-SPOTenantCdnEnabled -CdnType Public



