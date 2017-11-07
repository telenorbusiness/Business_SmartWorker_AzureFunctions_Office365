var Promise = require('bluebird');
var requestPromise = require('request-promise');

module.exports = function (context, req) {
        Promise
            .try(() =>  {
                return auth(req, context);
            })
            .then((graphToken) => {
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
            .catch(authHeaderUndefinedError,(error) => {
                let res = {
                    status: 403,
                    body: error+""
                };
                return context.done(null, res);
            })
            .catch((error) => {
                context.log(error);
                let res = {
                    body: error.message
                };
                return context.done(null, res);
            });
};

function auth(req, context) {

    if (typeof (req.headers.authorization) === 'undefined') {
        throw new authHeaderUndefinedError('Auth header is undefined');
    }

    var guidToken = req.headers.authorization.replace("Bearer ", "");
    var requestOptions = {
        method: 'POST',
        resolveWithFullResponse: true,
        json: true,
        simple: false,
        uri: getEnvironmentVariable("validatePartnerEndpoint"), //Using dev for now. Prod one is in env variables
        headers: {
            'Authorization': 'Basic ' + getEnvironmentVariable("clientIdSecret")
        },
        body: {
            "token": guidToken
        }
    };

    return requestPromise(requestOptions)
        .then(function (response) {
            if (response.statusCode === 200 && typeof(response.body.error) === "undefined") {
                return response.body.azureUserToken
                //return the azure graph token when ready
            }
            else {
                throw new atWorkValidateError(JSON.stringify(response.body), response.statusCode);
            }
        });
}

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
        "search": {
            "type": "local",
            "placeholder": "Søk etter dokumenter"
        },
        "sections": [
            {
            "header": 'Dokumenter',
            "searchableParameters" : ["title"],
            "rows": []
            }
        ],
    };

    for (let i = 0; i < documents.value.length; i++) {
        if(!documents.value[i].folder) {
            microApp.sections[0].rows.push(
                {
                "type": "text",
                "title": documents.value[i].name,
                "onClick": {
                "type": "open-url",
                "url": documents.value[i].webUrl
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

class authHeaderUndefinedError extends Error {}
class atWorkValidateError extends Error {
    constructor(message, response) {
        super(message);
        this.response = response;
    }
}
