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
        throw new atWorkValidateError("Atwork validation error", response);      }
    })
    .then(appointments => {
      let res = {
        body: createMicroApp(appointments, context)
      };
      idplog({message: "Completed sucessfully", sender: "calendar_tile", status: "200"})
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      let res = {
        status: error.response.status,
        body: error.response.message
      };
      idplog({message: "Error: atWorkValidateError: "+error, sender: "calendar_microapp", status: error.response.status})
      return context.done(null, res);
    })
    .catch(error => {
      context.log(error);
      let res = {
        status: 500,
        body: "An unexpected error occurred"
      };
      idplog({message: "Error: unknown error: "+error, sender: "calendar_microapp", status: "500"})
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
    .add(2, "months")
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
        maxDate +
        "&$select=subject,start,end,responseStatus,bodyPreview,location&$orderby=start/dateTime&$top=20"
    ),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions)
    .then(response => {
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

function createMicroApp(appointments, context) {
  let notRespondedRows = [];
  let respondedRows = [];
  let sections = [];
  let sectionIndex = -1;
  let lastRespondedDay = "";
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
          url: "https://outlook.office.com/owa/?path=/mail/inbox"
        }
      });
    } else {
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
        numContentLines: 1
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
  microApp.sections.push({
    rows: [
      {
        type: "button",
        title: "VIS KALENDER",
        onClick: {
          type: "open-url",
          url: "https://outlook.office365.com/owa/?path=/calendar/view/Month"
        }
      }
    ]
  });
  return microApp;
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

class atWorkValidateError extends Error {
  constructor(message, response) {
    super(message);
    this.response = response;
  }
}