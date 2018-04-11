# Business_SmartWorker_AzureFunctions_Office365
Azure functions for creating MicroApps and Tiles integrating with office365

__URL TO DEPLOY TEMPLATE__
 https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Ftelenorbusiness%2FBusiness_SmartWorker_AzureFunctions_Office365%2Fmaster%2Fazuredeploy.json

 __HOW TO UPDATE YOUR FUNCTION APPS__
  In order to update your function apps whenever there is changes in this repo, you have to do the following in the azure portal: Function apps -> "yourfunctionapp" -> Platform features -> Deployment options -> Sync.

__Changelog__

  11-04-2018. Added the "SCM_USE_FUNCPACK" app setting to bundle npm modules resulting in better performance.
