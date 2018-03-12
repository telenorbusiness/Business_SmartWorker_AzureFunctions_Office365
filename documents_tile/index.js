var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");

module.exports = function(context, req) {
  let graphToken;
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && response.azureUserToken) {
        graphToken = response.azureUserToken;
        return getDocumentsFromSharepoint(response.azureToken, context);
      } else {
        throw new atWorkValidateError(
          JSON.stringify(response.message),
          response.status
        );
      }
    })
    .catch(sharePointError, error => {
      context.log(error);
      context.log(
        "Error fetching from sharePoint, falling back to fetch from shared documents in onedrive"
      );
      return getDocuments(graphToken, context);
    })
    .then(documents => {
      let res = {
        body: createTile(documents)
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

function getDocuments(graphToken, context) {
  var requestOptions = {
    method: "GET",
    resolveWithFullResponse: true,
    json: true,
    simple: false,
    uri: "https://graph.microsoft.com/v1.0/me/drive/sharedWithMe",
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions).then(function(response) {
    if (response.statusCode === 200) {
      return response.body.value;
    } else {
      throw new Error(
        "Fetching documents returned with status code: " +
          response.statusCode +
          " and message: " +
          response.body.error.message
      );
    }
  });
}

function getDocumentsFromSharepoint(graphToken) {
  let hostName = getEnvironmentVariable("sharepointHostName");
  let relativePathName = getEnvironmentVariable("sharepointRelativePathName");

  if (!hostName || !relativePathName) {
    throw new sharePointError("Sharepoint env vars not set");
  }

  var requestOptions = {
    method: "GET",
    json: true,
    simple: true,
    uri: encodeURI(
      "https://graph.microsoft.com/v1.0/sites/" +
        hostName +
        ":/sites/" +
        relativePathName
    ),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions)
    .then(function(body) {
      let siteId = body.id;
      requestOptions.uri = encodeURI(
        "https://graph.microsoft.com/v1.0/sites/" +
          siteId +
          "/drive/root/children"
      );
      return requestPromise(requestOptions);
    })
    .then(function(response) {
      return response.value;
    })
    .catch(function(error) {
      throw new sharePointError(error);
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
    tile.footnote =
      getPrettyDate(lastModifiedDoc.lastModifiedDateTime) +
      ", " +
      lastModifiedDoc.lastModifiedBy.user.displayName;
  }

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

class sharePointError extends Error {
  constructor(message) {
    super(message);
  }
}
