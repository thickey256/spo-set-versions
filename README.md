# spo-set-versions
PowerShell script to Major and Minor versions across SharePoint sites and remove out-of-date versions

This script can run in one of two modes
1. What If - Non destructive dry run, will report on potential space savings
2. Delete Old Versions - Same as above but will actually delete versions in SharePoint

## Creating a Certificate
1. Edit the CreateSelfSignedCert.ps1 file.
2. Replace *{certificateName}* with something like *SharePoint Version Control*.
3. Navigate to directory in PowerShell and run ./CreateSelfSignedCert.ps1
4. Make a note of the random string that gets output (the Thumbprint).

## How to setup as an Azure app
1. Sign in to the [Azure portal](https://portal.azure.com/)
2. Navigate to App Registrations
3. New Registration

* Name your app something like "Sharepoint Version Control"
* Select the correct account type (normally Single Tenant)
* Leave Redirect URL empty
* Click Register

* Make note of the Application (client) ID
* Make note of the Directory (tenant) ID
* Click "View API permissions"

### Sharepoint Version Control | API permissions
* Click Add a permission
* Click the Sharepoint block
* Click Application permissions
* Check 'Sites.FullControl.All' or 'Sites.Selected' if you only want to look at one SharePoint Site
* Click Add permissions button
* You should see your new permission for SharePoint Listed
* Click 'Grant admin consent'
![Successful Permission!](/readme_images/app-setup-3.png)

### Upload Certificate
* upload it

## ParametersVersions.json (the config)
<details>
	<summary>How to setup the config file</summary>
	
1. Replace the *TenantID Value* with the *Directory (tenant) ID* you noted down in the stage above (line 3).
2. Replace the *clientId Value* with the *Application (client) ID* you noted down in the stage above (line 6).
3. Replace the *thumbprint Value* with the *Thumbprint* string you noted down in step 4 of Creating a Certificate (line 11).
4. Change *majorVersionCount* to the number of major version you want to keep, or leave at the default value (line 19)
5. Change *minorVersionCount* to the number of minor version you want to keep, or leave at the default value (line 22)
6. Change *whatIfMode Value* to *true*
7. Change *deleteOldMajorVersions Value* to *false*

*Steps 4 and 5 are WELL worth double checking, probably best to do a try run first before deleting all your files!*
</details>

## SitesToProcess.csv
This file holds the SharePoint URLs for the sites you want to process.


## Powershell
1. Install [PNP module](https://www.powershellgallery.com/packages/PnP.PowerShell/1.12.8-nightly) (used so PowerShell can talk to Sharepoint)
` Install-Module -Name PnP.PowerShell`
2. Make sure PowerShell has script execution permissions, [follow instructions here](https://learn.microsoft.com/en-gb/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.3).