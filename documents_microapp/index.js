var Promise = require('bluebird');
var requestPromise = require('request-promise');
const reftokenAuth = require('../auth');

module.exports = function (context, req) {
        Promise
            .try(() =>  {
                return reftokenAuth(req);
            })
            .then((response) => {
                if(response.status === 200 && response.azureUserToken) {
                    return getDocuments(response.azureToken, context);
                }
                else {
                    throw new atWorkValidateError(response.message, response.status);
                }
            })
            .then((documents) => {
                context.res = {
                    body: createMicroApp(documents)
                };
                return context.done(null, res);
            })
            .catch(atWorkValidateError,(error) => {
                context.log("Logger: "+error.response);
                context.res = {
                    status: error.response,
                    body: JSON.parse(error.message)
                }
                return context.done(null, res);
            })
            .catch((error) => {
                context.log(error);
                context.res = {
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
        uri: 'https://graph.microsoft.com/beta/me/drive/root/children',
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
        })
}

function createMicroApp(documents) {
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
            rows: []
            }
        ],
    };

    for (let i = 0; i < documents.value.length; i++) {
        if(!documents.value[i].folder) {
            microApp.sections[0].rows.push(
                {
                type: "text",
                title: documents.value[i].name,
                onClick: {
                type: "open-url",
                url: documents.value[i].webUrl
                }
            });
        }
    }
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
