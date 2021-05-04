/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
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
  "{email}":emailRegex,
  "{user}":userRegex,
  "{date}":dateRegex,
  "{interval}":intervalRegex,
  "{newLine}":newLineRegex,
  "{url}":urlRegex
}

// TODO: parcurgerea inversa
// documentatie
// detectare links

function escapeRegex(string) {
  return string.replace(/[-\/\\^$*+?.()|[\]]/g, '\\$&');
}

// Array that stores the start position of each `{}` field.
var indexes = [];

// Returns special fields defined inside the `regexString` parameter that were found inside `bodyContent`.
function bodyContains(bodyContent, regexString) {
  escapeRegex(regexString);
  // Structure to keep order of found fields.
  let foundFields = [];
  // Replace globally, case insensitive.
  var finalRegex = regexString.replace(/{email}|{user}|{date}|{interval}|{newLine}|{url}/gi, function(foundField) {
    foundFields.push(foundField);
    // `regex.source` returns a string literal that doesn't contain delimiters, thus allowing easier
    // regex concatenation. 
    indexes.push(regexString.indexOf(foundField) - 1);
    return regexMap[foundField].source;
  });

  var match = bodyContent.match(finalRegex)[0];

  var returnFields = [];

  // Each field's index will change based on the length of the found string `tempMatch`.
  var toAdd = 0;
  foundFields.forEach(function (entry) {
    // Start searching from indexes determined before.
    indexes[0] = indexes[0] + toAdd;
    var subMatch = match.substring(indexes[0]);
    indexes.shift();
    var tempMatch = subMatch.match(regexMap[entry])[0];
    toAdd += tempMatch.length - entry.length;
    returnFields.push(tempMatch);
  });
  return returnFields;
}

const baseURL = "https://google.com";
var something = [];
function processMessage(arg) {
  var messageFromDialog = JSON.parse(arg.message);
  something.push(messageFromDialog);
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
  if(queryParameters) {
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

function runUrlPath(startingRequest) {
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

var requestPaths = {
  "Cerere": [
    new Request("https://localhost:3000/link.html", "GET", {
      "name":"test",
      "password":"bop"
    }),
    new Request("https://localhost:3000/link.html", "GET", {
      "password":"test"
    })
  ]
}

export async function run() {
  // Get a reference to the current message
  var item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
  item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        
    } else {
      document.getElementById("item-subject").innerHTML = "<b>Code:</b> <br/>" + something.size;
      var content = asyncResult.value.trim();
      document.getElementById("item-body").innerHTML = "<b>Body:</b> <br/>" + content;
      var contentTest = "Perioada solicitata / The requested period : test@yahoo.com dadada 11-05-2000 - 11-08-2000\nRares";
      var test = "Perioada solicitata / The requested period : {email} dadada {interval}{newLine}{user}";
      
      var returnedFields = bodyContains(contentTest, test);
      if(returnedFields.size != 0) {
        var htmlContent = "<b>Fields:</b> <br/>";
        returnedFields.forEach(function (entry) {
          htmlContent += '- ' + entry + "<br/>";
        });
        document.getElementById("item-test").innerHTML = htmlContent;
      }
      document.getElementById("item-link").onclick = getStatus;
      
    }
 });  
}
