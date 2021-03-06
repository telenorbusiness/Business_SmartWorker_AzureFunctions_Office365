var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var idplog = require("../logging");

module.exports = function(context, req) {
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && response.azureUserToken) {
        return getAppointments(context, response.azureUserToken);
      } else {
        throw new atWorkValidateError("Atwork validation error", response);
      }
    })
    .then(appointments => {
      let res = {
        body: createTile(appointments)
      };
      idplog({message: "Completed sucessfully", sender: "calendar_tile", status: "200"})
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      let res = {
        status: error.response.status,
        body: error.response.message
      };
      idplog({message: "Error: atWorkValidateError: "+error, sender: "calendar_tile", status: error.response.status})
      return context.done(null, res);
    })
    .catch(error => {
      context.log(error);
      let res = {
        status: 500,
        body: "An unexpected error occurred"
      };
      idplog({message: "Error: unknown error: "+error, sender: "calendar_tile", status: "500"})
      return context.done(null, res);
    });
};

function getAppointments(context, graphToken) {
  const now = moment()
    .tz("Europe/Oslo")
    .utc()
    .format("YYYY-MM-DDTHH:mm:ss");
  const maxDate = moment()
    .utc()
    .add(2, "months")
    .format("YYYY-MM-DD");
  var requestOptions = {
    method: "GET",
    resolveWithFullResponse: true,
    json: true,
    simple: false,
    uri: encodeURI(
      "https://graph.microsoft.com/beta/me/calendarview?startdatetime=" +
        now +
        "&enddatetime=" +
        maxDate +
        "&$select=subject,start,end,responseStatus&$orderby=start/dateTime&$top=20"
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
      return null;
    })
    .catch(error => {
      context.log("Graph api call returned with error: " + error.statusCode);
      return null;
    });
}

function createTile(appointments) {
  if (appointments === null) {
    return { type: "text", text: "Feil", subtext: "ved henting av kalender" };
  }

  let tile = {};
  tile.type = "text";
  tile.text = "Ingen avtaler";
  tile.subtext = "";
  tile.footnote = "";
  tile.notifications = 0;
  tile.onClick = {
    type: "micro-app",
    apiUrl:
      "https://" +
      getEnvironmentVariable("appName") +
      ".azurewebsites.net/api/calendar_microapp"
  };

  let appointmentAdded = false;
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
      tile.notifications++;
      continue;
    } else if (!appointmentAdded) {
      tile.text = getPrettyDate(appointments[i].start.dateTime);
      tile.subtext =
        "kl " +
        getPrettyTime(appointments[i].start.dateTime) +
        " - " +
        getPrettyTime(appointments[i].end.dateTime);
      tile.footnote = appointments[i].subject;
      appointmentAdded = true;
    }
  }

  return tile;
}

function getPrettyDate(date) {
  let time = moment
    .utc(date)
    .tz("Europe/Oslo")
    .locale("nb");
  if (time.isSame(new Date(), "day")) {
    return "Idag";
  } else {
    return time.format("ddd Do MMM");
  }
}

function getPrettyTime(date) {
  let time = moment
    .utc(date)
    .tz("Europe/Oslo")
    .locale("nb");

  return time.format("H:mm");
}

function getEnvironmentVariable(name) {
  return process.env[name];
}

class atWorkValidateError extends Error {
  constructor(message, response) {
    super(message);
    this.response = response;
  }
}
