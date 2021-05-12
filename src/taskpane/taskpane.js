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

const emailRegex = /(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))/;

// Matches only if the the string contains a single word.
const userRegex = /\b[A-Z].*?\b/;
const split = / - /;
const newLineRegex = /(\r\n|\r|\n)/;
// Matches every date format (hyphen, slash & dot).
const dateRegex = /([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0]?[1-9]|[1][0-2])[./-]([0-9]{4})/;

const intervalRegex = new RegExp(dateRegex.source + split.source + dateRegex.source);

// Matches URLs and returns the following array structure:
// Index 0: Whole URL,    Index 1: Protocol, Index 2: Host, Index 3: Path
// Index 4: Query string, Index 5: Hash mark
const urlRegex = /([a-z]+\:\/+)([^\/\s]+)([a-z0-9\-@\^=%&;\/~\+]*)[\?]?([^ \#\r\n]*)#?([^ \#\r\n]*)/mig;

// Structure used to replace each string on the left with it's corresponding regex.
const regexMap = {
  "{email}": emailRegex,
  "{user}": userRegex,
  "{date}": dateRegex,
  "{interval}": intervalRegex,
  "{newLine}": newLineRegex,
  "{url}": urlRegex
}

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

// Opens the given URL in a nonmodal window.
function openEmbedded(url) {
  var dialog;
  Office.context.ui.displayDialogAsync(url,
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

const paramSpecialRegex = /(request)([0-9])([a-zA-Z]+)=([a-zA-Z0-9]+)/;
const paramFieldRegex = /([a-zA-Z0-9]+)=({[a-zA-Z]})/;

let requestPaths = []
let foundFields = []

// Extracts query params from the given URL.
function getQuery(url) {
  var query = [],
    href = url || window.location.href;

  href.replace(/[?&](.+?)=([^&#]*)/g, function (_, key, value) {
    query.push(key + '=' + decodeURI(value).replace(/\+/g, ' '));
  });

  return query;
}

// Returns the full query string after processing the special query params defined in the json file.
function constructQueryString(request, requestPos, pathIdentification) {
  let queryString = '';
  request.params.forEach(function (param) {
    let match = param.match(paramSpecialRegex);
    let fieldMatch = param.match(paramFieldRegex);
    if (match != null) {
      let requestNo = match[2];
      let paramLocation = match[3];
      let paramName = match[4];
      // Find values in the body of previously sent requests.
      if (paramLocation == 'body') { // e.g. `request0body=name`
        if (requestNo < requestPos && requestNo >= 0) {
          let bodyResponse = requestPaths[pathIdentification].requests[requestNo].response;
        }
      
      // Find param of previously sent requests.
      } else if (paramLocation == 'param') { // e.g.  `request0body=name`
        let paramsOtherRequest = requestPaths[pathIdentification].requests[requestNo].params;
        paramsOtherRequest.forEach(function (param2) {
          if(param2.match(paramName + '=')) {
            queryString += param2 + '&';
          }
        });
      }
    } else if(fieldMatch != null) {
      // Get the value of the special field found in the email subject/body.
      returnFields.forEach(function (element) {
        if(element['key'] == fieldMatch[2]) {
          queryString += fieldMatch[1] + element['value'] + '&';
        }
      })
    } else {
      queryString += param + '&';
    }
  });
  return queryString;
}

// Outputs requests' statuses.
function logRequest(url, status, content) {
  if(document.getElementById('debugOn').checked) {
    var temp = '';
    if(status == 200) {
      temp = content + '<div style="color:green"' +  '>'  + url + ' ' + status + '</div>';
    } else {
      temp = content + '<div style="color:red"' +  '>' + url + ' ' + status + '</div>';
    };
    document.getElementById("item-log").innerHTML = temp.toString();
  }
}

// Recursive function that sends HTTP requests for the given request chain.
function followRequestPath(pathIdentification, requestNo, ev) {
  var xmlHttp = new XMLHttpRequest();
  let request = requestPaths[pathIdentification].requests[requestNo];
  if (requestNo == requestPaths[pathIdentification].requests.length) {
    Office.context.mailbox.item.notificationMessages.addAsync("Info", {
      type: "informationalMessage",
      message: "Request path completed successfully.",
      icon: "iconid",
      persistent: false
    })
    return;
  }
  xmlHttp.onreadystatechange = function () {
    if (xmlHttp.readyState == 4 && xmlHttp.status == 200) {
      if (requestNo >= 0 && requestNo < requestPaths[pathIdentification].requests.length) {
        requestPaths[pathIdentification].requests[requestNo].response = xmlHttp.response;
        console.log('Success' + ' 200');
        ++requestNo;
        followRequestPath(pathIdentification, requestNo);
        
        logRequest(request.url, xmlHttp.status, document.getElementById('item-log').innerHTML);
      }
    } else if (xmlHttp.readyState == 4 && xmlHttp.status != 200) {
      Office.context.mailbox.item.notificationMessages.addAsync("Error", {
        type: "informationalMessage",
        message: "Request path execution failed.",
        icon: "iconid",
        persistent: false
      })
      logRequest(request.url, xmlHttp.status, document.getElementById('item-log').innerHTML);
    }
  }
  let queryString = constructQueryString(request, requestNo, pathIdentification);

  // Append the query params to the URL if the request type is `GET`.
  if(request.type.localeCompare('GET') != 1) {
    xmlHttp.open(request.type, request.url + queryString, /* async */ true);
    xmlHttp.send(null);
  // Append the query params in the request body otherwise.
  } else {
    xmlHttp.open(request.type, request.url, /* async */ true);
    xmlHttp.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
    xmlHttp.send(queryString);
  }
}

// Returns a randomly generated string composed of letters & digits.
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
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      let body = asyncResult.value.trim() + item.subject;
      document.getElementById("item-body").innerHTML = "<b>Body:</b> <br/>" + body;
      document.getElementById("item-log").innerHTML = "<b>Log:</b> <br/>";
      
      // Concatenate subject and email body into a single string.
      let content = subject + body;

      // Loop through each email template.
      for (const [key, value] of Object.entries(patterns.patterns[0])) {
        foundFields = bodyContains(contentTest, key);
        if (foundFields != null) {
          var htmlContent = "<b>Fields:</b> <br/>";
          foundFields.forEach(function (element) {
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
            // Action URLs defined in the email pattern.
            let urls;
            var urlContent = "<b>Actions:</b> <br/>";
            for (const [actionRegex, requestArray] of Object.entries(value.actions[0])) {
              urls = getFieldValue(bodyContains(body, actionRegex), '{url}');
              urls.forEach(function (url) {
                // Generate a string that can be used to identify request paths for each URL found in the body.
                let pathIdentification = makeid(10);
                urlContent += '<div id=' + pathIdentification + '> - ' + url + "</div> <br/>";
                document.getElementById("item-actions").innerHTML = urlContent;
                requestPaths[pathIdentification] = requestArray;
                // Add starting URL to the requestPath.
                let startingUrlParams = getQuery(url);
                
                requestPaths[pathIdentification].requests.unshift({
                  'url': url,
                  'type': 'GET',
                  'params': startingUrlParams
                });
                if(document.getElementById('sendRequests').checked) {
                  document.getElementById(pathIdentification).addEventListener("click",
                    followRequestPath.bind(null, pathIdentification, 0), true);
                } else {
                  // Open embedded browser on the given URL.
                  document.getElementById(pathIdentification).addEventListener("click",
                    openEmbedded.bind(null, url), true);
                }
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
    } else {
      item.notificationMessages.addAsync("Error", {
        type: "informationalMessage",
        message: "Email body acquisition failed.",
        icon: "iconid",
        persistent: false
      })
    }
  });
}
