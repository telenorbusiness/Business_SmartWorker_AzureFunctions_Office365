# Business_SmartWorker_AzureFunctions_Office365
Azure functions for creating MicroApps and Tiles integrating with office365

__URL TO DEPLOY TEMPLATE__
 https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Ftelenorbusiness%2FBusiness_SmartWorker_AzureFunctions_Office365%2Fmaster%2Fazuredeploy.json

 __HOW TO UPDATE YOUR FUNCTION APPS__
  In order to update your function apps whenever there is changes in this repo, you have to do the following in the azure portal: Function apps -> "yourfunctionapp" -> Platform features -> Deployment options -> Sync.

__Changelog__

  __05-07-2018.__ Added configuration microapp and new configuration endpoint. Users now get their sharepoint id in the documents microapp from a config ID (old way will still be used as fallback for now).
  __11-04-2018.__ Added the "SCM_USE_FUNCPACK=1" app setting to bundle npm modules resulting in better performance. If you have deployed before this date you need to
  add this setting manually in application settings(Function apps -> "yourfunctionapp" -> Application settings). Add new setting with the name SCM_USE_FUNCPACK and set the value to 1, save and then update your function app(see instructions above).
