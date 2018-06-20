var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var azure = require("azure-storage");
var tableService = azure.createTableService(getEnvironmentVariable("AzureWebJobsStorage"));

module.exports = function(context, req) {
  let graphToken;
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && response.azureUserToken) {
        graphToken = response.azureUserToken;
        let upn = getUpnFromJWT(graphToken, context);
        return getStorageInfo(upn, context);
      } else {
        throw new atWorkValidateError(
          JSON.stringify(response.message),
          response.status
        );
      }
    })
    .then((sharepointId) => {
      return getRecentActivity(graphToken, sharepointId);
    })
    .then(activities => {
      let res = {
        body: createTile(activities)
      };
      return context.done(null, res);
    })
    .catch(tableStorageError, error => {
      context.log("User not in storage, creating generic tile");
      let res = {
        body: createGenericTile()
      };
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      context.log("Logger: " + error.response);
      let res = {
        status: error.response,
        body: JSON.parse(error.message)
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

function getUpnFromJWT(azureToken, context) {
  let arrayOfStrings = azureToken.split(".");

  let userObject = JSON.parse(new Buffer(arrayOfStrings[1], "base64").toString());

  return userObject.upn.toLowerCase();
}

function getStorageInfo(rowKey, context) {
  return new Promise((resolve, reject) => {
    tableService.retrieveEntity("documents", "user_sharepointsites", rowKey, (err, result, response) => {
      if(!err) {
        resolve(result.sharepointId._);
      }
      else {
        context.log(err);
        reject(new tableStorageError(err));
      }
    });
  });
}

function getRecentActivity(graphToken, sharepointId) {
  const requestOptions = {
    method: "GET",
    json: true,
    simple: true,
    uri: encodeURI(
      "https://graph.microsoft.com/beta/sites/" +
      sharepointId +
      "/drive/activities?$expand=driveItem&$top=20"),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions)
    .then((response) => {
      return response.value;
    });
}

function createTile(activities = []) {
  var tile = {
    type: "icon",
    iconUrl:
      "https://api.smartansatt.telenor.no/cdn/office365/dokumenter.png",
    footnote: "Se delte dokumenter",
    onClick: {
      type: "micro-app",
      apiUrl:
        "https://" +
        getEnvironmentVariable("appName") +
        ".azurewebsites.net/api/documents_microapp"
    }
  };

  for (let i = 0; i < activities.length; i++) {
    const activity = activities[i];

    if(!activity.driveItem) {
      continue;
    }
    else if(activity.driveItem.file && (activity.action.edit || activity.action.create || activity.action.comment || activity.action.rename || activity.action.move)) {
      tile.footnote = "Siste endring: " + getPrettyDate(activity.times.recordedDateTime);
      return tile;
    }

  }

}

function createGenericTile() {
  var tile = {
    type: "icon",
    iconUrl: "https://api.smartansatt.telenor.no/cdn/office365/dokumenter.png",
    footnote: "Ikke registrert"
  };
  return tile;
}

function getPrettyDate(date) {
  let time = moment
    .utc(date)
    .tz("Europe/Oslo")
    .locale("nb");
  let now = moment
    .utc()
    .tz("Europe/Oslo")
    .locale("nb");
  if (time.isSame(now, "day")) {
    return time.format("H:mm");
  } else {
    return time.format("Do MMM");
  }
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

class tableStorageError extends Error {
  constructor(message) {
    super(message);
  }
}

