var Promise = require('bluebird');
var requestPromise = require('request-promise');
var moment = require('moment-timezone');
const reftokenAuth = require('../auth');

module.exports = function (context, req) {
        Promise
            .try(() =>  {
                return reftokenAuth(req);
            })
            .then((response) => {
                if(response.status === 200 && response.azureUserToken) {
                    return getMails(response.azureUserToken);
                }
                else {
                    throw new atWorkValidateError(response.message, response.status);
                }
            })
            .then((mails) => {
                context.res = {
                    body: createMicroApp(mails)
                };
                return context.done(null, res);
            })
            .catch(atWorkValidateError,(error) => {
                context.log("Atwork error: " + error.message);
                context.res = {
                    status: error.response,
                    body: error.message
                }
                return context.done(null, res);
            })
            .catch((error) => {
                context.log(error);
                context.res = {
                    body: error.message
                };
                return context.done(null, res);
            });
};

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
        id: "emails_main",
        sections: [
            {
            header: 'Ulest e-post',
            rows: []
            }
        ],
    };

    for (let i = 0; i < mails.value.length; i++) {
      if(!mails.value[i].isRead){
        microApp.sections[0].rows.push(
          {
            type: "text",
            title: mails.value[i].subject,
            subtitle: mails.value[i].sender.emailAddress.address,
            text: moment.utc(mails.value[i].sentDateTime).tz('Europe/Oslo').locale('nb').format("LLL"),
            onClick: {
              type: "open-url",
              url: mails.value[i].webLink
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
