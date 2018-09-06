var Promise = require("bluebird");
var azure = Promise.promisifyAll(require("azure-storage"));
const reftokenAuth = require("../auth");
const Joi = require('joi');


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
    .then(() => {
      return validateBody(req.body);
    })
    .then(response => {
      return insertUserInfo(req.body.configId, req.body.sharepointInfo, context);
    })
    .then(result => {
      let message =  "Oppdaterte sharepoint Id";
      let res = {
        body: { type: "reload", title: "Sharepoint konfigurasjon", message: message }
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
    .catch(tableStorageError, error => {
      context.log("Could not insert config: " + error);
      let res = {
        body: { type: "feedback", title: "Feil", message: "Feil ved opprettelse av konfigurasjon" }
      };
      return context.done(null, res);
    })
    .catch(validationError, error => {
      context.log("Error validating body: " + error);
      let res = {
        body: { type: "feedback", title: "Feil", message: "Feil ved validering av instilling" }
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

function insertUserInfo(configId, sharepointInfo, context) {
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
        sharepointInfo: entGen.String(JSON.stringify(sharepointInfo))
      };
      return tableService.insertOrReplaceEntityAsync("documents", entity);
    })
    .then(result => {
      context.log("Added row! -> " + JSON.stringify(result));
      return result;
    })
    .catch((error) => {
      throw new tableStorageError(error);
    });
}

function checkAuthKey(key) {
  if (key === process.env["configKey"]) {
    return true;
  } else throw new Error("Not a valid config key");
}

function validateBody(body) {

  return new Promise((resolve, reject) => {
    const schema = {
      configId: Joi.string().guid().required(),
      sharepointInfo: Joi.object().keys({
        displayName: Joi.string().required(),
        id: Joi.string().required(),
        webUrl: Joi.string().optional()
      }),
      configKey: Joi.string().guid().required()
    };

    Joi.validate(body, schema, (err, value) => {
      if(err === null) {
        resolve(true);
      }
      else {
        reject (new validationError(err));
      }
    });
  })
}

function getEnvironmentVariable(name) {
  return process.env[name];
}

class tableStorageError extends Error {
  constructor(message) {
    super(message);
  }
}

class validationError extends Error {
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