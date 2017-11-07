var Promise = require('bluebird');
var requestPromise = require('request-promise');
var moment = require('moment-timezone');

module.exports = function (context, req) {
        Promise
            .try(() =>  {
                return auth(req);
            })
            .then((graphToken) => {
                return getMails(graphToken);
            })
            .then((mails) => {
                let res = {
                    body: createMicroApp(mails)
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
                return response.body.azureUserToken
                //return the azure graph token when ready
            }
            else {
                throw new atWorkValidateError(JSON.stringify(response.body), response.statusCode);
            }
        });
}

function getMails(graphToken) {
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
                return response.body;
            }
            else {
                 throw new Error('Fetching emails returned with status code: ' + response.statusCode + " and message: " + response.body.error.message);
            }
        })
}

function createMicroApp(mails) {
    var microApp = {
        "sections": [
            {
            "header": 'Ulest e-post',
            "rows": []
            }
        ],
    };

    for (let i = 0; i < mails.value.length; i++) {
      if(!mails.value[i].isRead){
        microApp.sections[0].rows.push(
          {
            "type": "text",
            "title": mails.value[i].subject,
            "subtitle": mails.value[i].sender.emailAddress.address,
            "text": moment.utc(mails.value[i].sentDateTime).tz('Europe/Oslo').locale('nb').format("LLL"),
            "onClick": {
              "type": "open-url",
              "url": mails.value[i].webLink
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
