
var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var azure = require("azure-storage");
var lodash = require("lodash");
var tableService = azure.createTableService(getEnvironmentVariable("AzureWebJobsStorage"));

module.exports = function(context, req) {
//  let graphToken;
//  let sub;
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && !response.message) {

        if(checkIsAdminFromSmartAnsatt(response) && response.configId) {
          return getStorageInfo(response.configId, context)
          .then((defaultValue) => {
            const res = {
              body: createSearchMicroApp(response.configId, defaultValue, response.azureUserToken, req.query.search)
            };
            return context.done(null, res);
          })
        }
        else {
          const res = {
            body: {
              id: "denied_config",
              sections: [{
                rows: [{ type: "text", text: "Access denied. You're not an administrator."}]
              }]
            }
          };
          return context.done(null, res);
        }
        //graphToken = response.azureUserToken;
        //sub = getUpnFromJWT(graphToken, context);
        //return getStorageInfo(sub, context);
      } else {
        throw new atWorkValidateError("Atwork validation error", response);
      }
    })
    .catch(tableStorageError, error => {
      let res = {
        body: createEmptyMicroApp()
      };
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      let res = {
        status: error.response.status,
        body: error.response.message
      };
      return context.done(null, res);
    })
    .catch(error => {
      context.log(error);
      let res = {
        status: 500,
        body: "An unexpected error occurred"
      };
      return context.done(null, res);
    });
};

function getUpnFromJWT(azureToken, context) {
  let arrayOfStrings = azureToken.split(".");

  let userObject = JSON.parse(
    new Buffer(arrayOfStrings[1], "base64").toString()
  );

  return userObject.upn.toLowerCase();
}

function getStorageInfo(rowKey, context) {
  return new Promise((resolve, reject) => {
    tableService.retrieveEntity("documents", "user_sharepointsites", rowKey, (err, result, response) => {
      if(!err) {
        resolve(JSON.parse(result.sharepointInfo._));
      }
      else {
        if(err.statusCode === 404) {
          resolve(null);
        }
        else {
          reject(new tableStorageError(err));
          context.log(err);
        }
      }
    });
  });
}


function createMicroApp(configId, defaultValue) {

  var microApp = {
    id: "documents_config",
    sections: [
      {
        header: "Legg til sharepoint id",
        rows: [
        {
          type: "input",
          title: "Sharepoint ID",
          form: {
            type: "text",
            dataType: "text",
            inputKey: "sharepointId",
            inputPlaceholder: "Sharepoint ID",
            defaultValue: defaultValue === null ? "" : defaultValue //FETCH NÅVÆRENDE
          }
        },
        {
         type: "button",
         onClick: {
          type: "call-api",
          url: "https://" +getEnvironmentVariable("appName") +".azurewebsites.net/api/documents_config_new", //NEW ENDPOINT
          httpMethod: "PUT",
          httpBody: { configId: configId, configKey: process.env["configKey"] },
          includeInputKeys: ["sharepointId"]
          },
         title: "LAGRE"
        }]
      }]
  };

  return microApp;
}

function createSearchMicroApp(configId, defaultValue, graphToken, search = '') {
  const requestOptions = {
    method: "GET",
    json: true,
    simple: true,
    uri: encodeURI(
      "https://graph.microsoft.com/v1.0/sites/?search=" +
      search +
      "&$top=20"),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions)
    .then((response) => {
      const sharepointSites = response.value;

      var microApp = { id: "config_microapp_documents",
                        search: {
                          type: "server",
                          placeholder: "Søk etter sharepoint sider knyttet til din bedrift"
                        },
                        sections: []
                      };

      if(defaultValue !== null) {
        microApp.sections.push({
          header: "Nåværende sharepoint side",
          rows: [{
            type: "rich-text",
            title: defaultValue.displayName,
            text: defaultValue.webUrl,
            content: "Id: " + defaultValue.id
          }]
        });
      }

      var rows = [];

      sharepointSites.forEach((site) => {
        rows.push({
          type: "rich-text",
          title: site.displayName,
          text: site.webUrl,
          content: "Id: " + site.id,
          onClick: {
            type: "call-api",
            url: "https://" +getEnvironmentVariable("appName") +".azurewebsites.net/api/documents_config_new",
            httpMethod: "GET",
            httpBody: {
              configId: configId,
              sharepointInfo: {
                displayName: site.displayName,
                webUrl: site.webUrl,
                id: site.id
              }
            },
            alert: {
              type: "query",
              title: "Godkjenn valg",
              message: "Er du sikker på at du vil gi tilgang til " + site.displayName + " til de valgte gruppene?"
            }
          }
        });
      });

      microApp.sections.push({
        header: "Tilgjengelige sharepoint sider",
        rows: rows
      });

      return microApp;

    })
    .catch((error) => {
      return { id: "graphExplorer_error", sections: [{ rows: [ { type: "text", text: "Feil ved kommunikasjon mot Microsoft Graph API'et" }]}]};
    });
}

function createEmptyMicroApp() {
  var microApp = {
    id: "documents_empty",
    sections: [
      {
        rows: [
          {
            type: "rich-text",
            content: "Det er ingen sharepoint sider knyttet til din config.."
          }
        ]
      }
    ]
  };
  return microApp;
}


function checkIsAdminFromSmartAnsatt( res ){
  if( true === lodash.get(res, "administrator", false) ){
    return true;
  }
  else {
    return false;
  }
}

function getEnvironmentVariable(name) {
  return process.env[name];
}

class atWorkValidateError extends Error {
  constructor(message, response) {
    super(message);
    this.response = response;
  }
}

class tableStorageError extends Error {
  constructor(message) {
    super(message);
  }
}