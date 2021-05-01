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
// Matches every date format (hyphen, slash & dot).
let dateRegex = /([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0]?[1-9]|[1][0-2])[./-]([0-9]{4})/;

let intervalRegex = new RegExp(dateRegex.source + split.source + dateRegex.source);

// Structure used to replace each string on the left with it's corresponding regex.
let regexMap = {
  "{email}":emailRegex,
  "{user}":userRegex,
  "{date}":dateRegex,
  "{interval}":intervalRegex
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
  var finalRegex = regexString.replace(/{email}|{user}|{date}|{interval}/gi, function(foundField) {
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

const baseURL = "https://httpbin.org/get";

function getStatus() {
  const Http = new XMLHttpRequest();
  Http.open("GET", baseURL);
  Http.send();

  Http.onreadystatechange = (e) => {
    document.getElementById("item-subject").innerHTML = "<b>Code:</b> <br/>" + Http.status;
  }
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
      var contentTest = "Perioada solicitata / The requested period : test@yahoo.com dadada 11-05-2000 - 11-08-2000 Abc Rares";
      var test = "Perioada solicitata / The requested period : {email} dadada {interval} Abc {user}";
      
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
