var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var idplog = require("../logging");

module.exports = function(context, req) {
  const folderName = req.body.folderName;
  const depth = req.body.depth;
  const driveId = req.body.driveId;
  const itemId = req.body.itemId;
  Promise.try(() => {
    return reftokenAuth(req);
  })
  .then(response => {
    if (response.status === 200 && response.azureUserToken) {
      return getDocuments(response.azureUserToken, context, driveId, itemId);
    } else {
      throw new atWorkValidateError("Atwork validation error", response);      }
  })
  .then(documents => {
    let res = {
      body: {
        type: "response-view",
        view: createMicroApp(documents, folderName, depth)
      }
    };
    idplog({message: "Completed sucessfully", sender: "documents_microapp_subview", status: "200"});
    return context.done(null, res);
  })
  .catch(atWorkValidateError, error => {
    let res = {
      status: error.response.status,
      body: error.response.message
    };
    idplog({message: "Error: atWorkValidateError: "+error, sender: "documents_microapp_subview", status: error.response.status});
    return context.done(null, res);
  })
  .catch(error => {
    context.log(error);
    let res = {
      status: 500,
      body: "An unexpected error occurred"
    };
    idplog({message: "Error: unknown error: "+error, sender: "documents_microapp_subview", status: "500"});
    return context.done(null, res);
  });
};

function getDocuments(graphToken, context, driveId, itemId) {
  var requestOptions = {
    method: "GET",
    resolveWithFullResponse: true,
    json: true,
    simple: false,
    uri: encodeURI(
      "https://graph.microsoft.com/beta/drives/" +
        driveId +
        "/items/" +
        itemId +
        "/children"
    ),
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

function createMicroApp(documents, folderName, depth) {
  let fileRows = [];
  let folderRows = [];

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
            depth: depth++
          },
          httpMethod: "POST"
        }
      });
    }
  }

  var microApp = {
    id: "documents_subview_" + folderName + ":" + depth,
    search: {
      type: "local",
      placeholder: "Søk etter dokumenter"
    },
    sections: [
      {
        header: folderName,
        searchableParameters: ["title"],
        rows: folderRows.concat(fileRows)
      }
    ]
  };

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
