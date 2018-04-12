var requestPromise = require('request-promise');

module.exports = function(context, myTimer) {

  var requestOptions = {
    method: "GET",
    json: true,
    uri: encodeURI("https://" + process.env["appName"] + ".azurewebsites.net/api/documents_microapp")
  };

  return requestPromise(requestOptions)
    .then((res) => {
      context.log("Pinged documents with response: " + res);
      context.done();
    })
    .catch((err) => {
      context.log("Documents returned with: " + err);
      context.done();
    });

};
