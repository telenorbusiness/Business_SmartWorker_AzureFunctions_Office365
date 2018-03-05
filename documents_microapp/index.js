var Promise = require('bluebird');
var requestPromise = require('request-promise');
const reftokenAuth = require('../auth');

module.exports = function (context, req) {
    let graphToken;
        Promise
            .try(() =>  {
                return reftokenAuth(req);
            })
            .then((response) => {
                if(response.status === 200 && response.azureUserToken) {
                    graphToken = response.azureUserToken;
                    return getDocumentsFromSharepoint(response.azureToken, context);
                }
                else {
                    throw new atWorkValidateError(response.message, response.status);
                }
            })
            .catch(sharePointError, (error) => {
                context.log(error);
                context.log("Error fetching from sharePoint, falling back to fetch from shared documents in onedrive");
                return getDocuments(graphToken, context);
            })
            .then((documents) => {
                let res = {
                    body: createMicroApp(documents)
                };
                return context.done(null, res);
            })
            .catch(atWorkValidateError,(error) => {
                context.log("Logger: "+error.response);
                let res = {
                    status: error.response,
                    body: JSON.parse(error.message)
                }
                return context.done(null, res);
            })
            .catch((error) => {
                context.log(error);
                let res = {
                    status: 500,
                    body: "An unexpected error occurred"
                };
                return context.done(null, res);
            });
};

function getDocuments(graphToken, context) {
    var requestOptions = {
        method: 'GET',
        resolveWithFullResponse: true,
        json: true,
        simple: false,
        uri: 'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe',
        headers: {
            'Authorization': 'Bearer ' + graphToken
        },
    };

    return requestPromise(requestOptions)
        .then(function (response) {
            if(response.statusCode === 200) {
                return response.body;
            }
            else {
                throw new Error('Fetching documents returned with status code: ' + response.statusCode + " and message: " + response.body.error.message);
            }
        });
}

function getDocumentsFromSharepoint(graphToken) {
    let hostName = getEnvironmentVariable("sharepointHostName");
    let relativePathName = getEnvironmentVariable("sharepointRelativePathName");

    if(!hostName || !relativePathName) {
        throw new sharePointError('Sharepoint env vars not set');
    }

    var requestOptions = {
        method: 'GET',
        json: true,
        simple: true,
        uri: 'https://graph.microsoft.com/v1.0/sites/' + hostName + ':/sites/' + relativePathName,
        headers: {
            'Authorization': 'Bearer ' + graphToken
        },
    };

    return requestPromise(requestOptions)
        .then(function (body) {
            let siteId = body.id;
            requestOptions.uri = 'https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/root/children';
            return requestPromise(requestOptions);
        })
        .then(function(response) {
            return response;
        })
        .catch(function(error) {
            throw new sharePointError(error);
        });
}

function createMicroApp(documents) {

    let folderRows = [];
    let fileRows = [];
    for (let i = 0; i < documents.value.length; i++) {
        if(!documents.value[i].folder) {
            fileRows.push({
                type: "text",
                title: documents.value[i].name,
                onClick: {
                type: "open-url",
                url: documents.value[i].webUrl
                }
            });
        }
        else {
            let driveId;
            let itemId;
            // If driveitems are fetched with 'sharedWithMe'
            if(documents.value[i].remoteItem && documents.value[i].remoteItem.parentReference) {
                driveId = documents.value[i].remoteItem.parentReference.driveId;
                itemId = documents.value[i].remoteItem.id;
            }
            // If driveitems are fetched from sharePoint
            if(documents.value[i].id && documents.value[i].parentReference) {
                driveId = documents.value[i].parentReference.driveId;
                itemId = documents.value[i].id;
            }
            folderRows.push({
                type: "text",
                title: documents.value[i].name,
                onClick: {
                    type: "call-api",
                    url: "https://"+getEnvironmentVariable("appName")+".azurewebsites.net/api/documents_microapp_subview",
                    httpBody: {
                        folderName: documents.value[i].name,
                        driveId: driveId,
                        itemId: itemId,
                        depth: 0
                    },
                    httpMethod: "POST"
                }
            });
        }
    }

    var microApp = {
        id: "documents_main",
        search: {
            type: "local",
            placeholder: "Søk etter dokumenter"
        },
        sections: [
            {
            header: 'Dokumenter',
            searchableParameters : ["title"],
            rows: folderRows.concat(fileRows)
            }
        ],
    };

    return microApp;
}

function getEnvironmentVariable(name)
{
    return process.env[name];
}

class atWorkValidateError extends Error {
    constructor(message, response) {
        super(message);
        this.response = response;
    }
}

class sharePointError extends Error {
    constructor(message) {
        super(message);
    }
}
