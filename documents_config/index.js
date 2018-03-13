var Promise = require("bluebird");
var azure = Promise.promisifyAll(require("azure-storage"));

module.exports = function(context, req) {
  Promise.try(() => {
    return checkAuthKey(req.headers.authorization);
  })
    .then(response => {
      let sharepointId = req.body.sharepointId;
      let userId = req.body.upn;

      if (sharepointId && userId) {
        context.log("sharepointid and userId in body");
        return insertUserInfo(userId, sharepointId, context);
      }
      return null;
    })
    .then(result => {
      let res = {
        body:
          result === null
            ? "Missing necessary properties in body"
            : JSON.stringify(result)
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

function insertUserInfo(userId, sharepointId, context) {
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
        RowKey: entGen.String(userId),
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