var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var azure = Promise.promisifyAll(require("azure-storage"));

module.exports = function(context, req) {
  let graphToken;
  let sub;
  let appCreated = false;
  Promise.try(() => {
    context.log("Før ref token auth");
    return reftokenAuth(req);
  })
    .then(response => {
      context.log("Response from SA: " + JSON.stringify(response));
      if (response.status === 200 && response.azureUserToken) {
        graphToken = response.azureUserToken;
        sub = getUpnFromJWT(graphToken, context);
        return getStorageInfo(sub, context);
      } else {
        throw new atWorkValidateError(response.message, response.status);
      }
    })
    .then(sharepointId => {
      context.log("Before getDocumentsFromSharepoint with id: " + sharepointId);
      return getDocumentsFromSharepoint(graphToken, sharepointId);
    })
    .then(documents => {
      context.log(
        "Before createMicroApp with documents: " + JSON.stringify(documents)
      );
      let res = {
        body: createMicroApp(documents)
      };
      appCreated = true;
      return context.done(null, res);
    })
    .catch(tableStorageError, error => {
      context.log("Finding userSites returned with: " + error);
      let res = {
        body: createEmptyMicroApp()
      };
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      context.log("Logger: " + error);
      let res = {
        status: error.response,
        body: JSON.parse(error.message)
      };
      return context.done(null, res);
    })
    .catch(sharePointError, error => {
      context.log(error);
      context.log(
        "Error fetching from sharePoint, falling back to fetch from shared documents in onedrive"
      );
      let res = {
        status: 200,
        body: "Error from sharepoint"
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

  let userObject = JSON.parse(
    new Buffer(arrayOfStrings[1], "base64").toString()
  );

  return userObject.upn.toLowerCase();
}

function getStorageInfo(rowKey, context) {
  let tableService = azure.createTableService(
    getEnvironmentVariable("AzureWebJobsStorage")
  );
  return tableService
    .retrieveEntityAsync("documents", "user_sharepointsites", rowKey)
    .then(result => {
      return result.sharepointId._;
    })
    .catch(error => {
      throw new tableStorageError(error);
    });
}

function getSites(graphToken) {
  var requestOptions = {
    method: "GET",
    json: true,
    simple: false,
    uri: "https://graph.microsoft.com/beta/sites?search=",
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions).then(response => {
    return response.value;
  });
}

function getDocumentsFromSharepoint(graphToken, sharepointId) {
  var requestOptions = {
    method: "GET",
    json: true,
    simple: true,
    uri: encodeURI(
      "https://graph.microsoft.com/beta/sites/" +
        sharepointId +
        "/drive/root/children"
    ),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions)
    .then(function(response) {
      return response;
    })
    .catch(function(error) {
      throw new sharePointError(error);
    });
}

function createMicroApp(documents) {
  let folderRows = [];
  let fileRows = [];
  for (let i = 0; i < documents.value.length; i++) {
    if (!documents.value[i].folder) {
      fileRows.push({
        type: "rich-text",
        title: documents.value[i].name,
        tag: getPrettyDate(documents.value[i].lastModifiedDateTime),
        thumbnailUrl:
          "https://api.smartansatt.telenor.no/cdn/office365/files.png",
        onClick: {
          type: "open-url",
          url: documents.value[i].webUrl
        }
      });
    } else {
      let driveId;
      let itemId;
      // If driveitems are fetched with 'sharedWithMe'
      if (
        documents.value[i].remoteItem &&
        documents.value[i].remoteItem.parentReference
      ) {
        driveId = documents.value[i].remoteItem.parentReference.driveId;
        itemId = documents.value[i].remoteItem.id;
      }
      // If driveitems are fetched from sharePoint
      if (documents.value[i].id && documents.value[i].parentReference) {
        driveId = documents.value[i].parentReference.driveId;
        itemId = documents.value[i].id;
      }
      folderRows.push({
        type: "rich-text",
        title: documents.value[i].name,
        tag: getPrettyDate(documents.value[i].lastModifiedDateTime),
        thumbnailUrl:
          "https://api.smartansatt.telenor.no/cdn/office365/folder.png",
        onClick: {
          type: "call-api",
          url:
            "https://" +
            getEnvironmentVariable("appName") +
            ".azurewebsites.net/api/documents_microapp_subview",
          httpBody: {
            folderName: documents.value[i].name,
            driveId: driveId,
            itemId: itemId,
            depth: 0
          },
          httpMethod: "POST"
        }
      });
    }
  }

  var microApp = {
    id: "documents_main",
    search: {
      type: "local",
      placeholder: "Søk etter dokumenter"
    },
    sections: [
      {
        searchableParameters: ["title"],
        rows: folderRows.concat(fileRows)
      }
    ]
  };

  return microApp;
}

function createEmptyMicroApp() {
  var microApp = {
    id: "documents_empty",
    sections: [
      {
        rows: [
          {
            type: "rich-text",
            content: "Det er ingen sharepoint sider knyttet til din bruker"
          }
        ]
      }
    ]
  };
  return microApp;
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

class tableStorageError extends Error {
  constructor(message) {
    super(message);
  }
}