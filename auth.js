const request = require("request-promise");
const Issuer = require("openid-client").Issuer;
const jose = require("node-jose");
const lodash = require("lodash");

Issuer.defaultHttpOptions = { timeout: 25000, retries: 2, followRedirect: true };
let promiseReferenceTokenClient = null;

let authenticateReferenceToken = function(req, context) {
  if(promiseReferenceTokenClient === null) {
    promiseReferenceTokenClient = Issuer.discover(process.env["idpUrl"])
    .then(issuer => {
      let keyString = new Buffer(process.env["authKey"], "base64").toString();
      return jose.JWK.asKey(keyString, "json").then(key => {
        let client = new issuer.Client(
          {
            client_id: process.env["clientId_new"],
            client_secret: process.env["clientSecret_new"],
            userinfo_signed_response_alg: "HS256",
            userinfo_encrypted_response_alg: "RSA1_5",
            userinfo_encrypted_response_enc: "A128CBC-HS256",
            redirect_uris: []
          },
          key.keystore
        );
        client.CLOCK_TOLERANCE = 300;
        return client;
      });
    })
  }

  return promiseReferenceTokenClient.then(client => {
    if (
      lodash.trim(
        lodash
          .get(req, "headers.authorization", "")
          .replace(/(?:(B|b)earer )/, "")
      ) === ""
    ) {
      return Promise.resolve({ status: 401, message: "Auth header missing" });
    } else {
      const token = req.headers.authorization.replace(/(?:(B|b)earer )/, "");

      return client.userinfo(token).then(userinfo => {
        if (userinfo.success === true) {
          if (!userinfo.azureUserToken) {
            return Promise.resolve({
              status: 400,
              message: "Missing azuretoken"
            });
          }
          return Promise.resolve({
            status: 200,
            azureUserToken: userinfo.azureUserToken
          });
        } else if (!lodash.isUndefined(userinfo.error)) {
          return Promise.resolve({ status: 200, message: userinfo });
        } else {
          throw new Error(
            `This is not a reference token: ${JSON.stringify(userinfo)}`
          );
        }
      });
    }
  });
};

module.exports = authenticateReferenceToken;
