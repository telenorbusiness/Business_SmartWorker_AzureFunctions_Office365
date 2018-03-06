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
            if(response.status === 200 && response.azureUserToken) {
                return getAppointments(context, response.azureUserToken);
            }
            else {
                context.log("AtWork responded with: " + JSON.stringify(response));
                throw new atWorkValidateError(response.message, response.status);
            }
        })
        .then((appointments) => {
            let res = {
                body: createMicroApp(appointments)
            };
            return context.done(null, res);
        })
        .catch(atWorkValidateError,(error) => {
            let res = {
                status: error.response,
                body: error.message
            };
            return context.done(null, res);
        })
        .catch((error) => {
            context.log(error);
            let res = {
                status: 500,
                body: 'An unexpected error occurred'
            };
            return context.done(null, res);
        });
};

function getAppointments(context, graphToken) {
    var requestOptions = {
        method: 'GET',
        resolveWithFullResponse: true,
        json: true,
        simple: true,
        uri: encodeURI('https://graph.microsoft.com/beta/me/calendarview?startdatetime=' + moment().startOf('day').utc().format()
        + '&enddatetime=' + moment().add(7, 'days').utc().format()),
        headers: {
            'Authorization': 'Bearer ' + graphToken
        }
    };

    return requestPromise(requestOptions)
        .then((response) => {
             context.log("response: " + JSON.stringify(response.body));
            if(response.statusCode === 200 ) {
                return response.body.value;
            }
            return [];
        })
        .catch((error) => {
            context.log(error.statusCode + " status fra Graph API");
            throw new Error('Graph API feil');
        })
}

function createMicroApp(appointments) {
    let rows = [];

    let microApp = {
        id: "calendar_main",
        sections: []
    };
    if(appointments[0]) {
     var lastDay = moment.utc(appointments[0].start.dateTime).tz('Europe/Oslo').locale('nb').day();
    }
    for(let i = 0; i < appointments.length; i++) {
        let day = moment.utc(appointments[i].start.dateTime).tz('Europe/Oslo').locale('nb').day();

        if(rows.length !== 0 && lastDay !== day) {
            microApp.sections.push({
                header: getDay(lastDay),
                rows: rows
            });
            rows = [];
        }
        rows.push({
            type: "rich-text",
            title: appointments[i].subject,
            content: appointments[i].bodyPreview,
            tag: getPrettyDate(appointments[i].start.dateTime)
        });

        if(i === appointments.length - 1) {
            microApp.sections.push({
                header: getDay(day),
                rows: rows
            });
        }
        lastDay = moment.utc(appointments[i].start.dateTime).tz('Europe/Oslo').locale('nb').day();
    }

    if(microApp.sections.length === 0) {
        microApp.sections.push({
            rows: [{
                type: "text",
                title: "Ingen avtaler den neste uken"
            }]
        });
    }
    return microApp;
}

function getEnvironmentVariable(name)
{
    return process.env[name];
}

function getPrettyDate(callDateTime) {
    let time = moment.utc(callDateTime).tz('Europe/Oslo').locale('nb');
    if (time.isSame(new Date(), "day")) {
        return time.format('LT');
    }
    else {
        return time.format('LT');
    }
}

function getDay(dayNumber) {
    let today = moment.utc().tz('Europe/Oslo').locale('nb').day();

    switch(dayNumber) {
        case 0:
            return "Søndag";
        case 1:
            return "Mandag";
        case 2:
            return "Tirsdag";
        case 3:
            return "Onsdag";
        case 4:
            return "Torsdag";
        case 5:
            return "Fredag";
        case 6:
            return "Lørdag";
        default:
            return getPrettyDate(moment().day(dayNumber));
    }
}

class atWorkValidateError extends Error {
    constructor(message, response) {
        super(message);
        this.response = response;
    }
}
