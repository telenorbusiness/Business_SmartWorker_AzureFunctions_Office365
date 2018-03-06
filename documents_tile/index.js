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
                    let res = {
                    body: createTile()
                    };
                    return context.done(null, res);
                }
                else {
                    throw new atWorkValidateError(response.message, response.status);
                }
            })
            .catch(atWorkValidateError,(error) => {
                let res = {
                  status: error.response,
                  body: error.message
                }
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

function createTile() {
    var tile = {
        "type": "icon",
        "iconUrl": "https://cdn4.iconfinder.com/data/icons/universal-web-vol-1/256/43-512.png",
        "footnote": "Se delte dokumenter",
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

class atWorkValidateError extends Error {
    constructor(message, response) {
        super(message);
        this.response = response;
    }
}
