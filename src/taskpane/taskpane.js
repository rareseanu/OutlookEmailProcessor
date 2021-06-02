// JSON structure that stores email patterns.
let patterns = []

async function readJson() {
  let response = await fetch('/taskpane/patterns.json');
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
const dd_mm_yyyy = /([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0]?[1-9]|[1][0-2])[./-]([0-9]{4})/;
const mm_dd_yyyy = /([0]?[1-9]|[1][0-2])[./-]([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0-9]{4})/;
const yyyy_mm_dd = /([0-9]{4})[./-]([0]?[1-9]|[1][0-2])[./-]([0]?[1-9]|[1|2][0-9]|[3][0|1])/;
const dateRegex = new RegExp("(" + dd_mm_yyyy.source + "|" + mm_dd_yyyy.source + "|" + yyyy_mm_dd.source + ")");
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

function getDatesFromInterval(interval) {
  var dd_mm_yyyy_regex = new RegExp(dd_mm_yyyy.source, dd_mm_yyyy.flags + 'g');
  let dd_mm_yyyy_match = [...interval[0].matchAll(dd_mm_yyyy_regex)];
  if (dd_mm_yyyy_match != null) {
    let startDate = new Date(dd_mm_yyyy_match[0][3], dd_mm_yyyy_match[0][2] - 1, dd_mm_yyyy_match[0][1]);
    let endDate = new Date(dd_mm_yyyy_match[1][3], dd_mm_yyyy_match[1][2] - 1, dd_mm_yyyy_match[1][1]);
    return [startDate, endDate];
  }

  var mm_dd_yyyy_regex = new RegExp(mm_dd_yyyy.source, mm_dd_yyyy.flags + 'g');
  let mm_dd_yyyy_match = [...interval[0].matchAll(mm_dd_yyyy_regex)];
  if (mm_dd_yyyy_match != null) {
    let startDate = new Date(mm_dd_yyyy_match[0][3] - 1, mm_dd_yyyy_match[0][1], mm_dd_yyyy_match[0][2]);
    let endDate = new Date(mm_dd_yyyy_match[1][3] - 1, mm_dd_yyyy_match[1][1], mm_dd_yyyy_match[1][2]);
    return [startDate, endDate];
  }

  var yyyy_mm_dd_regex = new RegExp(yyyy_mm_dd.source, yyyy_mm_dd.flags + 'g');
  let yyyy_mm_dd_match = [...interval[0].matchAll(yyyy_mm_dd_regex)];
  if (yyyy_mm_dd_match != null) {
    let startDate = new Date(yyyy_mm_dd_match[0][1], yyyy_mm_dd_match[0][2] - 1, yyyy_mm_dd_match[0][3]);
    let endDate = new Date(yyyy_mm_dd_match[1][1], yyyy_mm_dd_match[1][2] - 1, yyyy_mm_dd_match[1][3]);
    return [startDate, endDate];
  }
}

function escapeRegex(string) {
  return string.replace(/[-\/\\^$*+?.()|[\]]/g, '\\$&');
}

// Stores field & value pairs under the following structure:
// e.g. `{key: "{fieldName}", value: "fieldValue"}`
var extractedFields;

function generateRegexFromPattern(regexString) {
  // Replace globally, case insensitive.
  var finalRegex = regexString.replace(/{email}|{user}|{date}|{interval}|{newLine}|{url}/gi, function (foundField) {
    return regexMap[foundField].source;
  });
  return finalRegex;
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
const paramFieldRegex = /([a-zA-Z0-9]+)=({[a-zA-Z]+})/;

let requestPaths = []

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

      // Find param of previously sent requests.
      if (paramLocation == 'param') { // e.g.  `request0param=name`
        let paramsOtherRequest = requestPaths[pathIdentification].requests[requestNo].params;
        paramsOtherRequest.forEach(function (param2) {
          if (param2.match(paramName + '=')) {
            queryString += param2 + '&';
          }
        });
      }
    } else if (fieldMatch != null) {
      // Get the value of the special field found in the email subject/body.
      extractedFields.forEach(function (element) {
        if (element['key'] == fieldMatch[2]) {
          queryString += fieldMatch[1] + '=' + element['value'] + '&';
        }
      })
    } else {
      queryString += param + '&';
    }
  });
  if (queryString != "") {
    queryString = queryString.slice(0, -1);
  }
  console.log(queryString);
  return queryString;
}

// Outputs requests' statuses.
function logRequest(url, status, content) {
  if (document.getElementById('debugOn').checked) {
    var temp = '';
    if (status == 200) {
      temp = content + '<div style="color:green"' + '>' + url + ' ' + status + '</div>';
    } else {
      temp = content + '<div style="color:red"' + '>' + url + ' ' + status + '</div>';
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
    });
    if (sendExcelPostRequest() != null) {
      Office.context.mailbox.item.notificationMessages.addAsync("Info", {
        type: "informationalMessage",
        message: "Excel action completed successfully.",
        icon: "iconid",
        persistent: false
      });
    }
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
  if (request.type.localeCompare('GET') != 1) {
    if (queryString != "") {
      queryString = '?' + queryString;
    }
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

function resetUI() {
  document.getElementById("item-subject").innerHTML = "";
  document.getElementById("item-log").innerHTML = "";
  document.getElementById("item-test").innerHTML = "";
  document.getElementById("title").innerText = "";
  document.getElementById("item-actions").innerHTML = "";
}

function sendExcelPostRequest() {
  let dates = getDatesFromInterval(getFieldValue(extractedFields, "{interval}"));
  let email = getFieldValue(extractedFields, "{email}")[0];
  if (dates != null && email != null) {
    var xhr = new XMLHttpRequest();

    xhr.open("POST", "https://localhost:3000/updateExcel", true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    let json = JSON.stringify({ dates: dates, email: email });
    xhr.send(json);
  }
}

async function run() {
  // Get a reference to the current message
  var item = Office.context.mailbox.item;
  // Write message property value to the task pane
  resetUI();
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
  item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      let body = asyncResult.value.trim();
      //document.getElementById("item-body").innerHTML = "<b>Body:</b> <br/>" + body;
      document.getElementById("item-log").innerHTML = "<b>Log:</b> <br/>";

      // Concatenate subject and email body into a single string.
      let content = item.subject + body;
      // Loop through each email template.
      for (const [key, value] of Object.entries(patterns.patterns[0])) {
        extractedFields = bodyContains(content, key);
        if (extractedFields != null) {
          var htmlContent = "<b>Fields:</b> <br/>";
          extractedFields.forEach(function (element) {
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
                if (document.getElementById('sendRequests').checked) {
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
          break;
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