var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var azure = require("azure-storage");
var tableService = azure.createTableService(getEnvironmentVariable("AzureWebJobsStorage"));

module.exports = function(context, req) {
  let graphToken;
  let sub;
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && response.azureUserToken) {
        graphToken = response.azureUserToken;
        sub = getUpnFromJWT(graphToken, context);
        return getStorageInfo(sub, context);
      } else {
        throw new atWorkValidateError(response.message, response.status);
      }
    })
    .then(sharepointId => {
      return {documents: getDocumentsFromSharepoint(graphToken, sharepointId), recentFiles: getRecentActivity(graphToken, sharepointId)};
    })
    .props()
    .then(({ documents, recentFiles }) => {
      let res = {
        body: createMicroApp(documents, recentFiles)
      };
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
      return response;
    })
    .catch((error) => {
      throw new sharePointError(error);
    });
}

function getActionType(action) {
  if(action.edit) {
    return "Redigert";
  }
  else if(action.create) {
    return "Opprettet";
  }
  else if(action.comment) {
    return "Ny kommentar";
  }
  else if(action.rename) {
    return "Nytt navn";
  }
  else if(action.move) {
    return "Flyttet";
  }
  else {
    return "";
  }
}

function createMicroApp(documents, activities) {
  let folderRows = [];
  let fileRows = [];
  let activityRows = [];
  let activitiesAdded = [];

  for (let i = 0; i < activities.value.length; i++) {
    const activity = activities.value[i];

    if(!activity.driveItem || activitiesAdded.includes(activity.driveItem.id)) {
      continue;
    }
    else if(activitiesAdded.length === 3) {
      break;
    }
    else if(activity.driveItem.file && (activity.action.edit || activity.action.create || activity.action.comment || activity.action.rename || activity.action.move)) {
      activityRows.push({
        type: "rich-text",
        title: activity.driveItem.name,
        content: getActionType(activity.action),
        tag: getPrettyDate(activity.times.recordedDateTime),
        thumbnailUrl:
          "https://api.smartansatt.telenor.no/cdn/office365/files.png",
        onClick: {
          type: "open-url",
          url: activity.driveItem.webUrl
        }
      });
      activitiesAdded.push(activity.driveItem.id);
    }

  }

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
        header: "Siste endringer",
        searchableParameters: ["title"],
        rows: activityRows
      },
      {
        header: "Bibliotek",
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