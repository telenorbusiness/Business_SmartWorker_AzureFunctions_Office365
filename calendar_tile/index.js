var Promise = require('bluebird');
var requestPromise = require('request-promise');
const reftokenAuth = require('../auth');
var moment = require('moment-timezone');

module.exports = function (context, req) {
    Promise
        .try(() =>  {
            return reftokenAuth(req);
        })
        .then((response) => {
            if(response.status === 200 && response.azureUserToken) {
                return getNumOfAppointments(context, response.azureUserToken);
            }
            else {
                throw new atWorkValidateError(JSON.stringify(response.message), response.status);
            }
        })
        .then((numberOfAppointments) => {
            context.res = {
                body: createTile(numberOfAppointments)
            };
            return context.done(null, res);
        })
        .catch(atWorkValidateError,(error) => {
            context.log("Atwork error. response " + JSON.stringify(error.response) );
            context.res = {
                status: error.response,
                body: JSON.parse(error.message)
            };
            return context.done(null, res);
        })
        .catch((error) => {
            context.log(error);
            context.res = {
                status: 500,
                body: 'An unexpected error occurred'
            };
            return context.done(null, res);
        });
};

function getNumOfAppointments(context, graphToken) {
    context.log('https://graph.microsoft.com/beta/me/calendarview?startdatetime=' + moment.utc().tz('Europe/Oslo').locale('nb').startOf('day').format()
        + '&enddatetime=' + moment().endOf('day').utc().format());
    var requestOptions = {
        method: 'GET',
        resolveWithFullResponse: true,
        json: true,
        simple: false,
        uri: 'https://graph.microsoft.com/beta/me/calendarview?startdatetime=' + moment().startOf('day').utc().format()
        + '&enddatetime=' + moment().endOf('day').utc().format(),
        headers: {
            'Authorization': 'Bearer ' + graphToken
        },
    };

    return requestPromise(requestOptions)
        .then((response) => {
            if(response.statusCode === 200 ) {
                return response.body.value.length;
            }
            return null;
        })
        .catch((error) => {
            context.log("Graph api call returned with error: " + error.statusCode);
            return null;
        })
}

function createTile(appointmentsToday) {
    let tile = {};
    let footNote = '';

    if(appointmentsToday === 0) {
        footNote = 'Ingen avtaler idag';
    }
    else if(appointmentsToday === 1) {
        footNote = '1 avtale idag';
    }
    else {
        footNote = appointmentsToday + ' avtaler idag';
    }
    if(appointmentsToday !== null) {
        tile = {
            "type": "icon",
            "iconUrl": "https://www.graphicsfuel.com/wp-content/uploads/2012/02/calendar-icon-512x512.png",
            "footnote": footNote,
            "title": "Kalender",
            "onClick": {
            "type": "micro-app",
            "apiUrl": "https://"+getEnvironmentVariable("appName")+".azurewebsites.net/api/calendar_microapp"
            }
        };
    }
    else {
        tile = {
            "type": "icon",
            "iconUrl": "https://www.graphicsfuel.com/wp-content/uploads/2012/02/calendar-icon-512x512.png",
            "footnote": "Kunne ikke hente kalender",
            "title": "Kalender"
        };
    }

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
