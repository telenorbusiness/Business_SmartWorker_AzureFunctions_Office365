var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var azure = require("azure-storage");
var tableService = azure.createTableService(process.env["AzureWebJobsStorage"]);

module.exports = function(context, myTimer) {
  context.done();
};
