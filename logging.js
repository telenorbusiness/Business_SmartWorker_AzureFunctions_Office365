let rp = require("request-promise");


function sendRequest({ method, uri, data, retry = true }){
  return rp({
    resolveWithFullResponse : true,
    json                    : true,
    simple                  : false,
    method,
    uri,
    headers                 : {
        Authorization: 'Basic ' + new Buffer(`${process.env["configKey"]}:${process.env["clientId_new"]}`).toString('base64')
    },
    body: ["post","put", "delete"].indexOf(method.toLowerCase()) !== -1 ? data : undefined
  })
  .then(response => {
    if( response.statusCode < 300 ){
        return response.body;
    }
    else if( retry === true && response.statusCode === 401 ){
        return sendRequest({ method, uri, data, retry: false });
    }
    else {
      throw new Error("Error while communicating with idp");
    }
  });
}

let idplog = function({ message, sender}){
  let idpUrl = process.env["idpUrl"].replace(/\.well-known\/openid-configuration/g, "");
  idpUrl = idpUrl+'echolog'

  const data = {
    sender,
    message
  }
  return sendRequest({ method: "put", uri: idpUrl, data})
};

module.exports = idplog;