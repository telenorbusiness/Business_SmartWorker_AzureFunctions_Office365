var Promise = require('bluebird');
var requestPromise = require('request-promise');

module.exports = function (context, req) {
        Promise
            .try(() =>  {
                return auth(req);
            })
            .then((response) => {
                let res = {
                    body: createTile()
                };
                return context.done(null, res);
            })
            .catch(atWorkValidateError,(error) => {
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

function auth(req) {
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
            }
            else {
               throw new atWorkValidateError(JSON.stringify(response.body), response.statusCode);
            }
        });
}

function createTile() {

    var tile = {
        "type": "icon",
        "iconUrl": "http://downloadicons.net/sites/default/files/business-document-icon-64269.png",
        "footnote": "Dokumenter",
        "onClick": {
        "type": "micro-app",
        "apiUrl": "https://"+getEnvironmentVariable("appName")+".azurewebsites.net/api/documents_microapp"
        }
    };
    return tile;
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
