var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var azure = Promise.promisifyAll(require("azure-storage"));

module.exports = function(context, req) {
  let graphToken;
  let sub;
  Promise.try(() => {
    context.log("Før ref token auth");
    return reftokenAuth(req);
  })
    .then(response => {
      context.log("Response from SA: " + JSON.stringify(response));
      if (response.status === 200 && response.azureUserToken && response.sub) {
        graphToken = response.azureUserToken;
        sub = response.sub;
        let sharepointId = req.query.sharepointId;
        if (sharepointId) {
          context.log("sharepointid in query");
          return insertUserInfo(sub, sharepointId, context);
        }
        context.log("Before getStorageInfo");
        return getStorageInfo(sub, context);
      } else {
        throw new atWorkValidateError(response.message, response.status);
      }
    })
    .then(sharepointId => {
      context.log("Before getDocumentsFromSharepoint with id: " + sharepointId);
      return getDocumentsFromSharepoint(graphToken, sharepointId);
    })
    .catch(sharePointError, error => {
      context.log(error);
      context.log(
        "Error fetching from sharePoint, falling back to fetch from shared documents in onedrive"
      );
      return getDocuments(graphToken, context);
    })
    .then(documents => {
      context.log("Before createMicroApp");
      let res = {
        body: createMicroApp(documents)
      };
      return context.done(null, res);
    })
    .catch(tableStorageError, error => {
      context.log("Finding userSites returned with: " + error);
      return getSites(graphToken);
    })
    .then(sharepointSites => {
      context.log("Before createSelectionMicroApp");
      let res = {
        body: createSelectionMicroApp(sharepointSites)
      };
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      context.log("Logger: " + error);
      let res = {
        status: error.response,
        body: JSON.parse(error)
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

function getStorageInfo(rowKey, context) {
    let tableService = azure.createTableService(
      getEnvironmentVariable("AzureWebJobsStorage")
    );
   return tableService.retrieveEntityAsync("userSites","user_sharepointsites",rowKey)
    .then((result) => {
      context.log("Response: " + JSON.stringify(result));
      return result.sharepointId;
    })
    .catch((error) => {
      throw new tableStorageError(error);
    });
  context.log("Kommer her");
}

function insertUserInfo(userId, sharepointId, context) {
  let tableService = azure.createTableService(
    getEnvironmentVariable("AzureWebJobsStorage")
  );
  let entGen = azure.TableUtilities.entityGenerator;

  return tableService.createTableIfNotExistsAsync("userSites")
    .then((response) => {
      context.log("Table created? ->" + JSON.stringify(response));
      let entity = {
        PartitionKey: entGen.String("user_sharepointsites"),
        RowKey: entGen.String(userId),
        sharepointId: entGen.String(sharepointId)
      };
      return tableService.insertEntityAsync("userSites", entity);
    })
    .then((result) => {
      context.log("Added row! -> " + JSON.stringify(result));
      return sharepointId;
    });
}

function getSites(graphToken) {
  var requestOptions = {
    method: "GET",
    json: true,
    simple: false,
    uri: "https://graph.microsoft.com/v1.0/sites?search=",
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions).then(response => {
    return response.value;
  });
}

function createSelectionMicroApp(siteRows = []) {
  const rows = siteRows.map(site => {
    let row = {
      type: "text",
      title: site.displayName,
      onClick: {
        type: "reload",
        queryParameters: { sharepointId: site.id }
      }
    };
    return row;
  });

  let microApp = {
    id: "documents_selection",
    search: {
      type: "local",
      placeholder: "Søk etter ditt område"
    },
    sections: [
      {
        header: "Velg ditt delte område",
        rows: rows
      }
    ]
  };
  return microApp;
}

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
      return response.body;
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

function getDocumentsFromSharepoint(graphToken, sharepointId) {
  //let hostName = getEnvironmentVariable("sharepointHostName");
  //let relativePathName = getEnvironmentVariable("sharepointRelativePathName");

  /* if (!hostName || !relativePathName) {
    throw new sharePointError("Sharepoint vars not set");
  }*/
  var requestOptions = {
    method: "GET",
    json: true,
    simple: true,
    uri: encodeURI(
      "https://graph.microsoft.com/v1.0/sites/" +
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
          "https://smartworker-dev-azure-api.pimdemo.no/microapps/random-static-files/icons/files.png",
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
          "https://smartworker-dev-azure-api.pimdemo.no/microapps/random-static-files/icons/folder.png",
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
