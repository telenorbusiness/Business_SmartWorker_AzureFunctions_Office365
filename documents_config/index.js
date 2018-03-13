var Promise = require("bluebird");
const reftokenAuth = require("../auth");
var azure = Promise.promisifyAll(require("azure-storage"));

module.exports = function(context, req) {

  Promise.try(() => {
    context.log("Før ref token auth");
    return reftokenAuth(req);
  })
    .then(response => {
      context.log("Response from SA: " + JSON.stringify(response));
      if (response.status === 200) {
        let sharepointId = req.body.sharepointId;
        let userId = req.body.upn;

        if (sharepointId && userId) {
          context.log("sharepointid and userId in query");
          return insertUserInfo(userId, sharepointId, context);
        }
        return null;
      } else {
        throw new atWorkValidateError(response.message, response.status);
      }
    })
    .then((res) => {
      let res = {
        body: res === null ? 'Missing necessary properties in body' : JSON.stringify(res)
      };
      return context.done(null, res);
    })
    .catch((error) => {
      let res = {
        status: 500,
        body: 'An unexpected error occurred'
      }
    });
  }

function insertUserInfo(userId, sharepointId, context) {
  let tableService = azure.createTableService(
    getEnvironmentVariable("AzureWebJobsStorage")
  );
  let entGen = azure.TableUtilities.entityGenerator;

  return tableService.createTableIfNotExistsAsync("documents_microapp")
    .then((response) => {
      context.log("Table created? ->" + JSON.stringify(response));
      let entity = {
        PartitionKey: entGen.String("user_sharepointsites"),
        RowKey: entGen.String(userId),
        sharepointId: entGen.String(sharepointId)
      };
      return tableService.insertOrReplaceEntityAsync("documents_microapp", entity);
    })
    .then((result) => {
      context.log("Added row! -> " + JSON.stringify(result));
      return result;
    });
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
