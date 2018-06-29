var Promise = require("bluebird");
var azure = Promise.promisifyAll(require("azure-storage"));
const reftokenAuth = require("../auth");


module.exports = function(context, req) {
  Promise.try(() => {
      return reftokenAuth(req);
  })
    .then(res => {
      if (res.status === 200 && !res.message) {
            return checkAuthKey(req.body.configKey);
      } else {
        throw new atWorkValidateError("AtWork validate error", res);
      }
    })
    .then(response => {
      let sharepointId = req.body.sharepointId;
      let configId = req.body.configId;

      if (sharepointId && configId) {
        return insertUserInfo(configId, sharepointId, context);
      }
      return null;
    })
    .then(result => {
      let message = result === null ? "Missing necessary properties in body" : "Oppdaterte sharepoint Id";
      let res = {
        body: { type: "feedback", title: "Sharepoint konfigurasjon", message: message }
      };
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      let res = {
        status: error.response.status,
        body: error.response.message
      };
      return context.done(null, res);
    })
    .catch(error => {
      let res = {
        status: 500,
        body: "An unexpected error occurred"
      };
      return context.done(null, res);
    });
};

function insertUserInfo(configId, sharepointId, context) {
  let tableService = azure.createTableService(
    getEnvironmentVariable("AzureWebJobsStorage")
  );
  let entGen = azure.TableUtilities.entityGenerator;

  return tableService
    .createTableIfNotExistsAsync("documents")
    .then(response => {
      context.log("Table created? ->" + JSON.stringify(response));
      let entity = {
        PartitionKey: entGen.String("user_sharepointsites"),
        RowKey: entGen.String(configId),
        sharepointId: entGen.String(sharepointId)
      };
      return tableService.insertOrReplaceEntityAsync("documents", entity);
    })
    .then(result => {
      context.log("Added row! -> " + JSON.stringify(result));
      return result;
    });
}

function checkAuthKey(key) {
  if (key === process.env["configKey"]) {
    return true;
  } else throw new Error("Not a valid config key");
}
function getEnvironmentVariable(name) {
  return process.env[name];
}

class tableStorageError extends Error {
  constructor(message) {
    super(message);
  }
}

class atWorkValidateError extends Error {
  constructor(message, response) {
    super(message);
    this.response = response;
  }
}