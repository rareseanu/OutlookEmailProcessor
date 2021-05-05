/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

// JSON structure that stores email patterns.
let patterns = []

async function readJson() {
  let response = await fetch('patterns.json');
  let patterns = await response.json();
  return patterns;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    // Load email patterns.
    readJson()
      .then(json => {
        patterns = json;
      })
  }
});

let emailRegex = /(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))/;

// Matches only if the the string contains a single word.
let userRegex = /\b[A-Z].*?\b/;
let split = / - /;
let newLineRegex = /(\r\n|\r|\n)/;
// Matches every date format (hyphen, slash & dot).
let dateRegex = /([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0]?[1-9]|[1][0-2])[./-]([0-9]{4})/;

let intervalRegex = new RegExp(dateRegex.source + split.source + dateRegex.source);

// Matches URLs and returns the following array structure:
// Index 0: Whole URL,    Index 1: Protocol, Index 2: Host, Index 3: Path
// Index 4: Query string, Index 5: Hash mark
let urlRegex = /([a-z]+\:\/+)([^\/\s]+)([a-z0-9\-@\^=%&;\/~\+]*)[\?]?([^ \#\r\n]*)#?([^ \#\r\n]*)/mig;

// Structure used to replace each string on the left with it's corresponding regex.
let regexMap = {
  "{email}": emailRegex,
  "{user}": userRegex,
  "{date}": dateRegex,
  "{interval}": intervalRegex,
  "{newLine}": newLineRegex,
  "{url}": urlRegex
}

// TODO: parcurgerea inversa
// documentatie
// detectare links

function escapeRegex(string) {
  return string.replace(/[-\/\\^$*+?.()|[\]]/g, '\\$&');
}

// Returns special fields defined inside the `regexString` parameter that were found inside `bodyContent`.
function bodyContains(bodyContent, regexString) {
  escapeRegex(regexString);
  // Array that stores the start position of each `{}` field.
  let indexes = [];

  // Structure to keep order of found fields.
  let foundFields = [];
  // Replace globally, case insensitive.
  var finalRegex = regexString.replace(/{email}|{user}|{date}|{interval}|{newLine}|{url}/gi, function (foundField) {
    foundFields.push(foundField);
    // `regex.source` returns a string literal that doesn't contain delimiters, thus allowing easier
    // regex concatenation. 
    indexes.push(regexString.indexOf(foundField) - 1);
    return regexMap[foundField].source;
  });

  var returnFields = [];
  let temp = bodyContent.match(finalRegex);
  if (!temp)
    return null;

  var match = temp[0];
  // Each field's index will change based on the length of the found string `tempMatch`.
  var toAdd = 0;
  foundFields.forEach(function (entry) {
    // Start searching from indexes determined before.
    indexes[0] = indexes[0] + toAdd;
    var subMatch = match.substring(indexes[0]);
    indexes.shift();
    var tempMatch = subMatch.match(regexMap[entry]);
    if (tempMatch != null) {
      var fullMatch = tempMatch[0];
      toAdd += fullMatch.length - entry.length;
      returnFields.push({ key: entry, value: fullMatch });
    }
  });
  return returnFields;
}

function getStatus() {
  var dialog;
  Office.context.ui.displayDialogAsync('https://localhost:3000/link.html',
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
        dialog = asyncResult.error.code;
      } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    });
}

function Request(url, requestType, queryParameters) {
  this.url = url;
  this.requestType = requestType;
  if (queryParameters) {
    let iterations = Object.keys(queryParameters).length;
    this.url += '?';
    for (const [key, value] of Object.entries(queryParameters)) {
      this.url += key + '=' + value;
      if (--iterations)
        this.url += '&';
    }
  }
  this.queryParameters = queryParameters;
  // Modify response after sending the request.
  this.reponse = null;
}

function Pattern(description, requestPath) {
  this.description = description;
  this.requestPaths = requestPath;
}

function runUrlPath(startingRequest, requestPath) {
  var dialog;
  Office.context.ui.displayDialogAsync(startingRequest.url,
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
        dialog = asyncResult.error.code;
      } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    });
}

// var requestPaths = {
//   "Cerere": [
//     new Request("https://localhost:3000/link.html", "GET", {
//       "name": "test",
//       "password": "bop"
//     }),
//     new Request("https://localhost:3000/link.html", "GET", {
//       "password": "test"
//     })
//   ]
// }

// Structure that stores email patterns 
var emailPatterns = {
  "Perioada solicitata / The requested period : {email} dadada {interval}{newLine}{user}":
    [
      new Pattern("Cerere Invoire", [
        new Request("https://localhost:3000/link.html", "GET", {
          "name": "test",
          "password": "bop"
        }),
        new Request("https://localhost:3000/link.html", "GET", {
          "password": "test"
        })
      ])
    ]
}

// Helper function that returns every value in the dictionary for the given key.
// It takes as parameter a dictionary whose structure is the following:
// e.g. `fieldDictionary = [ { key:{field}, value:{fieldValue}]`
function getFieldValue(fieldDictionary, key) {
  let values = []
  fieldDictionary.forEach(function (element) {
    if (element['key'] == key) {
      values.push(element['value'])
    }
  });
  return values;
}

let requestPaths = []

function followRequestPath(startingRequest) {
  // if(settings.isenabled follow programatically)
  let requestID = this.id;
  let requests = requestPaths[requestID].requests;
  console.log(typeof(requests));
  requests.forEach(function(request) {
    // Send request
    console.log(request);
  })
}

function makeid(length) {
  var result = [];
  var characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var charactersLength = characters.length;
  for (var i = 0; i < length; i++) {
    result.push(characters.charAt(Math.floor(Math.random() *
      charactersLength)));
  }
  return result.join('');
}

export async function run() {
  // Get a reference to the current message
  var item = Office.context.mailbox.item;
  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
  item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {

    } else {
      var content = asyncResult.value.trim();
      document.getElementById("item-body").innerHTML = "<b>Body:</b> <br/>" + content;
      var contentTest = "Perioada solicitata / The requested period : test@test.com dadada test@test2.com";

      // Loop through each email template.
      for (const [key, value] of Object.entries(patterns.patterns[0])) {
        var returnedFields = bodyContains(contentTest, key);
        if (returnedFields != null) {
          var htmlContent = "<b>Fields:</b> <br/>";
          returnedFields.forEach(function (element) {
            htmlContent += '- ' + element['key'] + ' : ' + element['value'] + "<br/>";
          })
          document.getElementById("item-test").innerHTML = htmlContent;
          document.getElementById("title").innerText = value.description;
          item.notificationMessages.addAsync("Info", {
            type: "informationalMessage",
            message: "Email pattern found: " + value.description,
            icon: "iconid",
            persistent: false
          })
          // Check if email pattern has a regex defined for URLs.
          if (value.actions != null) {
            var simulateUrl = "Pentru APROBARE accesati link-ul / For APPROVAL access the link :\nhttps://google.com ";
            // Action URLs defined in the email pattern.
            let urls;
            var urlContent = "<b>Actions:</b> <br/>";
            for (const [actionRegex, requestArray] of Object.entries(value.actions[0])) {
              urls = getFieldValue(bodyContains(simulateUrl, actionRegex), '{url}');
              urls.forEach(function (url) {
                // Generate a string that can be used to identify request paths for each URL found in the body.
                var pathIdentification = makeid(10);
                urlContent += '<div id=' + pathIdentification  + '> - ' + url + "</div> <br/>";
                document.getElementById("item-actions").innerHTML = urlContent;
                document.getElementById(pathIdentification).addEventListener("click", followRequestPath, false);
                requestPaths[pathIdentification] = requestArray;
                console.log(requestPaths);
              })
            }
          }
        }
      }
      if (document.getElementById("title").innerText == 'Process Email') {
        item.notificationMessages.addAsync("Info", {
          type: "informationalMessage",
          message: "No email pattern found.",
          icon: "iconid",
          persistent: false
        })
      }

    }
  });
}
