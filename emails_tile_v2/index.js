var Promise = require("bluebird");
var requestPromise = require("request-promise");
const reftokenAuth = require("../auth");
var moment = require("moment-timezone");

module.exports = function(context, req) {
  Promise.try(() => {
    return reftokenAuth(req);
  })
    .then(response => {
      if (response.status === 200 && response.azureUserToken) {
        return getUnreadMails(response.azureUserToken, context);
      } else {
        throw new atWorkValidateError(
          JSON.stringify(response.message),
          response.status
        );
      }
    })
    .then(unreadMails => {
      let res = {
        body: createTile(unreadMails)
      };
      return context.done(null, res);
    })
    .catch(atWorkValidateError, error => {
      context.log("Atwork error. response " + JSON.stringify(error.response));
      let res = {
        status: error.response,
        body: JSON.parse(error.message)
      };
      return context.done(null, res);
    })
    .catch(error => {
      context.log(error);
      let res = {
        body: error.message
      };
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
    resolveWithFullResponse: true,
    json: true,
    simple: false,
    uri: encodeURI(
      "https://graph.microsoft.com/beta/me/messages?$filter=isRead eq false and ReceivedDateTime ge " +
        dateOfLastEmail
    ),
    headers: {
      Authorization: "Bearer " + graphToken
    }
  };

  return requestPromise(requestOptions).then(function(response) {
    if (response.statusCode === 200) {
      return response.body.value;
    } else {
      context.log(
        "Fetching mails returned with status code: " + response.statusCode
      );
      return null;
    }
  });
}

function createTile(unreadMails) {
  if (unreadMails === null) {
    return { type: "text", text: "Feil", subtext: "ved henting av e-post" };
  }

  let tile = {};
  tile.type = "icon";
  tile.iconUrl =
    "https://api.smartansatt.telenor.no/cdn/office365/outlook.png";
  tile.notifications = 0;
  tile.onClick = {
    type: "open-url",
    url: "https://outlook.office.com/owa/?path=/mail/inbox"
  };

  for (let i = 0; i < unreadMails.length; i++) {
    tile.notifications++;
  }

  if (tile.notifications === 0) {
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