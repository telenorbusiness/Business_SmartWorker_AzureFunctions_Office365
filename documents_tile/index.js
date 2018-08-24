var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var azure = require("azure-storage");
var tableService = azure.createTableService(getEnvironmentVariable("AzureWebJobsStorage"));
var idplog = require("../logging");

module.exports = function(context, req) {
  let graphToken;
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && response.azureUserToken) {
        graphToken = response.azureUserToken;

        if(response.configId && response.configId !== null) {
          return getStorageInfo(response.configId, true, context);
        }
        else {
          let upn = getUpnFromJWT(graphToken);
          return getStorageInfo(upn, false, context);
        }
      } else {
        throw new atWorkValidateError("Atwork validation error", response);
      }
    })
    .then((sharepointId) => {
      return getRecentActivity(graphToken, sharepointId);
    })
    .then(activities => {
      let res = {
        body: createTile(activities)
      };
      idplog({message: "Completed sucessfully", sender: "document-tile", status: "200"})
      return context.done(null, res);
    })
    .catch(tableStorageError, error => {
      idplog({message: "Error: User not in storage, creating generic tile", sender: "document-tile", status: "204"})
      context.log("User not in storage, creating generic tile");
      let res = {
        body: createGenericTile()
      };
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      idplog({message: "Error: atWorkValidateError: "+error, sender: "document-tile", status: error.response.status})
      let res = {
        status: error.response.status,
        body: error.response.message
      };
      return context.done(null, res);
    })
    .catch(error => {
      idplog({message: "Error: Unknown error: "+error, sender: "document-tile", status: "500"})
      context.log(error);
      let res = {
        status: 500,
        body: "An unexpected error occurred"
      };
      return context.done(null, res);
    });
};

function getUpnFromJWT(azureToken) {
  let arrayOfStrings = azureToken.split(".");

  let userObject = JSON.parse(new Buffer(arrayOfStrings[1], "base64").toString());

  return userObject.upn.toLowerCase();
}

function getStorageInfo(rowKey, isConfigId, context) {
  return new Promise((resolve, reject) => {
    tableService.retrieveEntity("documents", "user_sharepointsites", rowKey, (err, result, response) => {
      if(!err) {
        if(!isConfigId) {
          resolve(result.sharepointId._);
        }
        else {
          let sharepointInfo = JSON.parse(result.sharepointInfo._);
          resolve(sharepointInfo.id);
        }
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

  return tile;

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

