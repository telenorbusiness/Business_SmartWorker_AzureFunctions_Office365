var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var azure = require("azure-storage");
var lodash = require("lodash");
var idplog = require("../logging");

var tableService = azure.createTableService(
  getEnvironmentVariable("AzureWebJobsStorage")
);

module.exports = function(context, req) {
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && !response.message) {
        if(!response.azureUserToken) {
          const res = {
            body: {
              id: "noAzureUserToken",
              sections: [
                {
                  rows: [
                    {
                      type: "text",
                      text: "Access denied. You do not have an Azure User Token."
                    }
                  ]
                }
              ]
            }
          };
          idplog({message: "NoAzureUserTokenError", sender: "documents_config_microapp", status: 200});
          return context.done(null, res);
        }
        else if (checkIsAdminFromSmartAnsatt(response) && response.configId) {
          return getStorageInfo(response.configId, context)
            .then(defaultValue => {
              return createSearchMicroApp(
                response.configId,
                defaultValue,
                response.azureUserToken,
                req.query.search
              );
            })
            .then(resolvedMicroApp => {
              const res = {
                body: resolvedMicroApp
              };
              return context.done(null, res);
            });
        } else {
          const res = {
            body: {
              id: "denied_config",
              sections: [
                {
                  rows: [
                    {
                      type: "text",
                      text: "Access denied. You're not an administrator."
                    }
                  ]
                }
              ]
            }
          };
          idplog({message: "Completed sucessfully", sender: "documents_config_microapp", status: "200"});
          return context.done(null, res);
        }
      } else {
        throw new atWorkValidateError("Atwork validation error", response);
      }
    })
    .catch(tableStorageError, error => {
      let res = {
        body: createEmptyMicroApp()
      };
      idplog({message: "tableStorageError: "+error, sender: "documents_config_microapp", status: "204"});
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      let res = {
        status: error.response.status,
        body: error.response.message
      };
      idplog({message: "error: atWorkValidateError: "+error, sender: "documents_config_microapp", status: error.response.status});
      return context.done(null, res);
    })
    .catch(error => {
      context.log(error);
      let res = {
        status: 500,
        body: "An unexpected error occurred"
      };
      idplog({message: "error: unknown error: "+error, sender: "documents_config_microapp", status: "500"});
      return context.done(null, res);
    });
};


function getStorageInfo(rowKey, context) {
  return new Promise((resolve, reject) => {
    tableService.retrieveEntity(
      "documents",
      "user_sharepointsites",
      rowKey,
      (err, result, response) => {
        if (!err) {
          if(result.sharepointInfo) {
            resolve(JSON.parse(result.sharepointInfo._));
          }
          else {
            resolve(null);
          }
        } else {
          if (err.statusCode === 404) {
            resolve(null);
          } else {
            reject(new tableStorageError(err));
            context.log(err);
          }
        }
      }
    );
  });
}

function createSearchMicroApp(
  configId,
  defaultValue = null,
  graphToken,
  search = ""
) {
  const requestOptions = {
    method: "GET",
    json: true,
    simple: true,
    uri: encodeURI(
      "https://graph.microsoft.com/v1.0/sites?search=" + search + "&$top=20"
    ),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions)
    .then(response => {
      const sharepointSites = response.value;

      var microApp = {
        id: "config_microapp_documents",
        search: {
          type: "server",
          placeholder: "Søk etter sharepoint sider knyttet til din bedrift"
        },
        sections: []
      };

      if (defaultValue !== null) {
        microApp.sections.push({
          header: "Nåværende sharepoint side",
          rows: [
            {
              type: "rich-text",
              title: defaultValue.displayName,
              text: defaultValue.webUrl
            }
          ]
        });
      }

      var rows = [];

      sharepointSites.forEach(site => {
        rows.push({
          type: "rich-text",
          title: site.displayName,
          text: site.webUrl,
          onClick: {
            type: "call-api",
            url:
              "https://" +
              getEnvironmentVariable("appName") +
              ".azurewebsites.net/api/documents_config_new",
            httpMethod: "POST",
            httpBody: {
              configId: configId,
              sharepointInfo: {
                displayName: site.displayName,
                webUrl: site.webUrl,
                id: site.id
              },
            configKey: process.env["configKey"]
            },
            alert: {
              type: "query",
              title: "Godkjenn valg",
              message:
                "Er du sikker på at du vil gi tilgang til " +
                site.displayName +
                " til de valgte gruppene?"
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
    .catch(error => {
      idplog({message: "Error: "+error, sender: "documents_config_microapp", status: error.statusCode});
      return {
        id: "graphExplorer_error",
        sections: [
          {
            rows: [
              {
                type: "text",
                text: "Feil ved kommunikasjon mot Microsoft Graph API'et"
              }
            ]
          }
        ]
      };
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

function checkIsAdminFromSmartAnsatt(res) {
  if (true === lodash.get(res, "administrator", false)) {
    return true;
  } else {
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
