var Promise = require('bluebird');
var requestPromise = require('request-promise');
var reftokenAuth = require('../auth');
var moment = require('moment-timezone');

module.exports = function (context, req) {

    Promise
        .try(() =>  {
           return reftokenAuth(req);
        })
        .then((response) => {
            if(response.statusCode === 200 && response.azureUserToken) {
                return getAppointments(context, response.azureUserToken);
            }
            else {
                throw new atWorkValidateError(response.message, response.status);
            }
        })
        .then((appointments) => {
            context.res = {
                body: createMicroApp(appointments)
            };
            return context.done(null, res);
        })
        .catch(atWorkValidateError,(error) => {
            context.res = {
                status: error.response,
                body: error.message
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

function getAppointments(context, graphToken) {
    context.log('https://graph.microsoft.com/beta/me/calendarview?startdatetime=' + moment.utc().tz('Europe/Oslo').locale('nb').startOf('day').format()
        + '&enddatetime=' + moment().endOf('day').utc().format());
    var requestOptions = {
        method: 'GET',
        resolveWithFullResponse: true,
        json: true,
        simple: false,
        uri: 'https://graph.microsoft.com/beta/me/calendarview?startdatetime=' + moment().startOf('day').utc().format()
        + '&enddatetime=' + moment().add(7, 'days').utc().format(),
        headers: {
            'Authorization': 'Bearer ' + graphToken;
        },
    };

    return requestPromise(requestOptions)
        .then((response) => {
            if(response.statusCode === 200 ) {
                return response.body.value;
            }
            return null;
        })
        .catch((error) => {
            //context.log('error fetching events from calendar api');
            context.log(error.statusCode + "status");
            return null;
        })
}

function createMicroApp(appointments) {
    let rows = [];
    rows.push({
        type: "text",
        title: "Test"
    });
    let microApp = {
        id: "calendar_main",
        sections: [
            {
                header: "Kalender",
                rows: rows
            }
        ]
    };

/*
    appointments.forEach(appointment => {
        let day =
        rows.push({
            type: "text",
            tit
        })
    })
*/
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
