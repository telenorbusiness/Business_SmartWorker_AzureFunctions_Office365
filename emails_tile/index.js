var Promise = require('bluebird');
var requestPromise = require('request-promise');

module.exports = function (context, req) {
    Promise
        .try(() =>  {
            return auth(req, context);
        })
        .then((graphToken) => {
            return getNumOfUnreadMails(graphToken, context);
        })
        .then((numOfUnreadMails) => {
            let res = {
                body: createTile(numOfUnreadMails)
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
                return response.body.azureUserToken;
                //return the azure graph token when ready
            }
            else {
                throw new atWorkValidateError(JSON.stringify(response.body), response.statusCode);
            }
        });
}

function getNumOfUnreadMails(graphToken, context) {
    var requestOptions = {
        method: 'GET',
        resolveWithFullResponse: true,
        json: true,
        simple: false,
        uri: 'https://graph.microsoft.com/beta/me/messages',
        headers: {
            'Authorization': 'Bearer ' + graphToken
        },
    };

    return requestPromise(requestOptions)
        .then(function (response) {
            if(response.statusCode === 200) {
                var numOfUnreadMails=0;
                for(let i = 0; i < response.body.value.length; i++) {
                    if(!response.body.value[i].isRead) {
                        numOfUnreadMails++;
                    }
                }
                return numOfUnreadMails;
            }
            else {
                throw new Error('Fetching mails returned with status code: ' + response.statusCode);
            }
        })
}

function createTile(numOfUnreadMails) {

    var tile = {
        "type": "icon",
        "iconUrl": "https://i.pinimg.com/736x/5d/a6/54/5da6545d8d46ea858c3507e914b0c027--email-icon-personality-types.jpg",
        "footnote": "Ulest e-post",
        "notifications": numOfUnreadMails,
        "onClick": {
            "type": "micro-app",
            "apiUrl": "https://"+getEnvironmentVariable("appName")+".azurewebsites.net/api/emails_microapp"
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
