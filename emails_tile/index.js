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
                return getNumOfUnreadMails(response.azureUserToken, context);
            }
            else {
                throw new atWorkValidateError(response.message, response.status);
            }
        })
        .then((numOfUnreadMails) => {
            context.res = {
                body: createTile(numOfUnreadMails)
            };
            return context.done(null, res);
        })
        .catch(atWorkValidateError,(error) => {
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
        type: "icon",
        iconUrl: "https://i.pinimg.com/736x/5d/a6/54/5da6545d8d46ea858c3507e914b0c027--email-icon-personality-types.jpg",
        footnote: "Ulest e-post",
        notifications: numOfUnreadMails,
        onClick: {
            type: "micro-app",
            apiUrl: "https://"+getEnvironmentVariable("appName")+".azurewebsites.net/api/emails_microapp"
        }
    };

    return tile;
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
