var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");
var idplog = require("../logging");

module.exports = function(context, req) {
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && response.azureUserToken) {
        return getUnreadMails(response.azureUserToken, context);
      } else {
        throw new atWorkValidateError("Atwork validation error", response);
      }
    })
    .then(unreadMails => {
      let res = {
        body: createTile(unreadMails)
      };
      idplog({message: "Completed sucessfully", sender: "email_tile_v2", status: "200"})
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      let res = {
        status: error.response.status,
        body: error.response.message
      };
      idplog({message: "Error: atWorkValidateError: "+error, sender: "email_tile_v2", status: error.response.status})
      return context.done(null, res);
    })
    .catch(error => {
      context.log(error);
      let res = {
        body: error.message
      };
      idplog({message: "Error: unknown error: "+error, sender: "email_tile_v2", status: "500"})
      return context.done(null, res);
    });
};

function getUnreadMails(graphToken, context) {
  let dateOfLastEmail = moment
    .utc()
    .tz("Europe/Oslo")
    .locale("nb")
    .subtract(14, "days")
    .format("YYYY-MM-DD");
  var requestOptions = {
    method: "GET",
    json: true,
    uri: encodeURI(
      "https://graph.microsoft.com/beta/me/mailFolders"
    ),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions)
    .then(function(response) {
      let inbox = response.value.find(folder => folder.wellKnownName === 'inbox');
      requestOptions.json = false;
      requestOptions.uri = encodeURI(
        "https://graph.microsoft.com/beta/me/mailFolders/" +
          inbox.id +
          "/messages/$count?$filter=isRead eq false and ReceivedDateTime ge " +
          dateOfLastEmail
      );
      return requestPromise(requestOptions);
    })
    .catch(error => {
      context.log(
        "Fetching mails returned with: " + error.StatusCodeError
      );
      return null;
    });
}

function createTile(unreadMails) {

  let tile = {};
  tile.type = "icon";
  tile.iconUrl = "https://api.smartansatt.telenor.no/cdn/office365/outlook.png";
  tile.notifications = unreadMails !== null ? parseInt(unreadMails) : 0;
  tile.onClick = {
    type: "open-url",
    url: "https://outlook.office.com/owa/?path=/mail/inbox"
  };

  if (unreadMails === null) {
    tile.footnote = "Gå til innboks";
  } else if (tile.notifications === 0) {
    tile.footnote = "Du har ingen uleste e-post";
  } else if (tile.notifications === 1) {
    tile.footnote = "Du har 1 ulest e-post";
  } else {
    tile.footnote = "Du har " + tile.notifications + " uleste e-post";
  }

  return tile;
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
