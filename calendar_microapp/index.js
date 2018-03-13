var Promise = require("bluebird");
var requestPromise = require("request-promise");
var reftokenAuth = require("../auth");
var moment = require("moment-timezone");

module.exports = function(context, req) {
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && response.azureUserToken) {
        return getAppointments(context, response.azureUserToken);
      } else {
        context.log("AtWork responded with: " + JSON.stringify(response));
        throw new atWorkValidateError(response.message, response.status);
      }
    })
    .then(appointments => {
      let res = {
        body: createMicroApp(appointments)
      };
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      let res = {
        status: error.response,
        body: error.message
      };
      return context.done(null, res);
    })
    .catch(error => {
      context.log(error);
      let res = {
        status: 500,
        body: "An unexpected error occurred"
      };
      return context.done(null, res);
    });
};

function getAppointments(context, graphToken) {
  const now = moment()
    .tz("Europe/Oslo")
    .utc()
    .format("YYYY-MM-DD");
  const maxDate = moment()
    .utc()
    .add(6, "months")
    .format("YYYY-MM-DD");
  var requestOptions = {
    method: "GET",
    resolveWithFullResponse: true,
    json: true,
    simple: true,
    uri: encodeURI(
      "https://graph.microsoft.com/beta/me/calendarview?startdatetime=" +
        now +
        "&enddatetime=" +
        maxDate
    ),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions)
    .then(response => {
      context.log("response: " + JSON.stringify(response.body));
      if (response.statusCode === 200) {
        return response.body.value;
      }
      return [];
    })
    .catch(error => {
      context.log(error.statusCode + " status fra Graph API");
      throw new Error("Graph API feil");
    });
}

function createMicroApp(appointments) {
  let notRespondedRows = [];
  let respondedRows = [];
  let sections = [];
  let sectionIndex = -1;
  let lastRespondedDay = "";
  let maxDate = moment()
    .add(14, "days")
    .utc()
    .tz("Europe/Oslo")
    .locale("nb");
  let now = moment
    .utc()
    .tz("Europe/Oslo")
    .locale("nb");

  for (let i = 0; i < appointments.length; i++) {
    let appointmentDate = moment
      .utc(appointments[i].start.dateTime)
      .tz("Europe/Oslo")
      .locale("nb");
    if (appointmentDate.isBefore(now)) {
      continue;
    } else if (appointments[i].responseStatus.response === "notResponded") {
      notRespondedRows.push({
        type: "rich-text",
        title: appointments[i].subject,
        text: getPrettyDate(appointmentDate),
        tag: getPrettyTime(appointmentDate),
        onClick: {
          type: "open-url",
          url: encodeURI(appointments[i].webLink)
        }
      });
    } else if (appointmentDate.isBefore(maxDate)) {
      if (!appointmentDate.isSame(lastRespondedDay, "day")) {
        sections.push({
          header: getPrettyDate(appointmentDate),
          rows: []
        });
        sectionIndex++;
      }
      sections[sectionIndex].rows.push({
        type: "rich-text",
        title: appointments[i].subject,
        text: appointments[i].location.displayName,
        content: appointments[i].bodyPreview,
        tag: getPrettyTime(appointmentDate),
        numContentLines: 1,
        onClick: {
          type: "open-url",
          url: encodeURI(appointments[i].webLink)
        }
      });
      lastRespondedDay = appointmentDate;
    }
  }

  if (notRespondedRows.length !== 0) {
    sections.unshift({
      header: "Invitasjoner",
      rows: notRespondedRows
    });
  }

  let microApp = {
    id: "calendar_main",
    sections: sections
  };

  if (microApp.sections.length === 0) {
    microApp.sections.push({
      rows: [
        {
          type: "text",
          title: "Ingen avtaler å vise"
        }
      ]
    });
  }
  return microApp;

  /*
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
    return microApp;*/
}

function getEnvironmentVariable(name) {
  return process.env[name];
}

function getPrettyDate(date) {
  let time = moment
    .utc(date)
    .tz("Europe/Oslo")
    .locale("nb");
  if (time.isSame(new Date(), "day")) {
    return "Idag";
  } else {
    return time.format("dddd Do MMM");
  }
}

function getPrettyTime(date) {
  let time = moment
    .utc(date)
    .tz("Europe/Oslo")
    .locale("nb");

  return time.format("H:mm");
}

function getDay(dayNumber) {
  let today = moment
    .utc()
    .tz("Europe/Oslo")
    .locale("nb")
    .day();

  switch (dayNumber) {
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
