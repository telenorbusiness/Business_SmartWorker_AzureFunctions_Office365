var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var azure = Promise.promisifyAll(require("azure-storage"));

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
      return getDocumentsFromSharepoint(graphToken, sharepointId);
    })
    .then(documents => {
      let res = {
        body: createTile(documents)
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
    let tableService = azure.createTableService(
      getEnvironmentVariable("AzureWebJobsStorage")
    );
   return tableService.retrieveEntityAsync("documents","user_sharepointsites",rowKey)
    .then((result) => {
      return result.sharepointId._;
    })
    .catch((error) => {
      throw new tableStorageError(error);
    });
}

function getDocumentsFromSharepoint(graphToken, siteId) {

  var requestOptions = {
    method: "GET",
    json: true,
    simple: true,
    uri: encodeURI(
        "https://graph.microsoft.com/beta/sites/" +
          siteId +
          "/drive/root/children"
      ),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions)
    .then(function(response) {
      return response.value;
    });
}

function createTile(documents = []) {
  var tile = {
    type: "icon",
    iconUrl:
      "https://smartworker-dev-azure-api.pimdemo.no/microapps/random-static-files/icons/dokumenter.png",
    footnote: "Se delte dokumenter",
    onClick: {
      type: "micro-app",
      apiUrl:
        "https://" +
        getEnvironmentVariable("appName") +
        ".azurewebsites.net/api/documents_microapp"
    }
  };

  if (documents.length > 0) {
    var lastModifiedDoc = documents[0];

    for (let i = 0; i < documents.length; i++) {
      if (
        moment(lastModifiedDoc.lastModifiedDateTime).isBefore(
          moment(documents[i].lastModifiedDateTime)
        )
      ) {
        lastModifiedDoc = documents[i];
      }
    }
    tile.footnote = "Siste endring: " + getPrettyDate(lastModifiedDoc.lastModifiedDateTime);
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

