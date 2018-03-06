# Business_SmartWorker_AzureFunctions_Office365
Azure functions for creating MicroApps and Tiles integrating with office365


<a href="https://azuredeploy.net/" target="_blank">
    <img src="http://azuredeploy.net/deploybutton.png"/>
</a>

#URL TO DEPLOY TEMPLATE
 https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Ftelenorbusiness%2FBusiness_SmartWorker_AzureFunctions_Office365%2Fdev%2Fazuredeploy.json

#SETTING UP SHAREPOINT IN DOCUMENTS MICRO APP
Once template is deployed, your documents micro app will default to showing documents available in the 'shared with me' drive of the user. In order to point to a specific sharepoint document library you need to set some environment variables. Enter the function app in the azure portal. Go to platform features -> application settings. Here you will need to add two new settings; "sharepointHostName" and "sharepointRelativePathName". Example, if the 'home' URL of your sharepoint site is "https://telenorsolutioncenter.sharepoint.com/sites/SA-testomrde", the "sharepointHostName" would be "telenorsolutioncenter.sharepoint.com" and the "sharepointRelativePathName" would be "SA-testomrde".
