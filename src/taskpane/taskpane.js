/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
var serviceRequest;
var xhr;
(() => {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
      $(document).ready(() => {
          //app.initialize();

          serviceRequest = new Object();
          serviceRequest.attachmentToken = "";
          serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
          serviceRequest.attachments = new Array();
      });
  };

  function initApp() {
      if (Office.context.mailbox.item.attachments == undefined) {
          showToast("Not supported", "Attachments are not supported by your Exchange server.");
      } else if (Office.context.mailbox.item.attachments.length == 0) {
          showToast("No attachments", "There are no attachments on this item.");
      } else {

          // Initalize a context object for the app.
          //   Set the fields that are used on the request
          //   object to default values.
          serviceRequest = new Object();
          serviceRequest.attachmentToken = "";
          serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
          serviceRequest.attachments = new Array();
      }
  };

})();

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("emailSubmit").onclick = submitForm;
  }
  initApp();

  serviceRequest = new Object();
  serviceRequest.attachmentToken = "";
  serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
  serviceRequest.attachments = new Array();

});

export async function submitForm() {
  var xmlhttp = new XMLHttpRequest();   // new HttpRequest instance 
  var theUrl = "https://localhost:44346/ExportToMpacs/Read";
  xmlhttp.open("POST", theUrl);
  xmlhttp.setRequestHeader("Content-Type", "application/json;charset=UTF-8");

  //getAttachmentToken();
  console.log(await GetMessageBody());
  testAttachments();
  //GetAttachments();
  // xmlhttp.send(JSON.stringify({
  //   a: Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text)
  // }));
}

const GetMessageBody = () => {
  return new Promise((resolve)=>{
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (res)=>{
      resolve(res.value);
    });
  });
}

function testAttachments() {
  Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
};

function attachmentTokenCallback(asyncResult, userContext) {
  if (asyncResult.status == "succeeded") {
      serviceRequest.attachmentToken = asyncResult.value;
      makeServiceRequest();
  }
  else {
      showToast("Error", "Could not get callback token: " + asyncResult.error.message);
  }
}

function showToast(title, message) {

  var notice = document.getElementById("notice");
  var output = document.getElementById('output');

  notice.innerHTML = title;
  output.innerHTML = message;

  $("#footer").show("slow");

  window.setTimeout(() => { $("#footer").hide("slow") }, 10000);
};

function makeServiceRequest() {
  var attachment;
  xhr = new XMLHttpRequest();

  // Update the URL to point to your service location.
  xhr.open("POST", "https://localhost:44346/ExportToMpacs", true);

  xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
  xhr.onreadystatechange = requestReadyStateChange;

  // Translate the attachment details into a form easily understood by WCF.
  for (var i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
      attachment = Office.context.mailbox.item.attachments[i];
      attachment = attachment._data$p$0 || attachment.$0_0;

      if (attachment !== undefined) {
          serviceRequest.attachments[i] = JSON.parse(JSON.stringify(attachment));
      }
  }

  // Send the request. The response is handled in the 
  // requestReadyStateChange function.
  xhr.send(JSON.stringify(serviceRequest));
};

function requestReadyStateChange() {
  if (xhr.readyState == 4) {
      if (xhr.status == 200) {
          var response = JSON.parse(xhr.responseText);
          if (!response.isError) {
              // The response indicates that the server recognized
              // the client identity and processed the request.
              // Show the response.
              var names = "<h2>Attachments processed: " + response.attachmentsProcessed + "</h2>";

              for (var i = 0; i < response.attachmentNames.length; i++) {
                  names += response.attachmentNames[i] + "<br />";
              }
              document.getElementById("names").innerHTML = names;
          } else {
              showToast("Runtime error", response.message);
          }
      } else {
          if (xhr.status == 404) {
              showToast("Service not found", "The app server could not be found.");
          } else {
              showToast("Unknown error", "There was an unexpected error: " + xhr.status + " -- " + xhr.statusText);
          }
      }
  }
};

export async function run() {
  
}

function GetSMTPAddress(mail) {
  return null;
}

function GetCurrentUserEmailAddress() {
  return Office.context.mailbox.userProfile.emailAddress;
}

function CleanEmailText(text) {
  // Microsofts version of characters
  text = text.Replace("\u2013", "-");
  text = text.Replace("\u2014", "-");
  text = text.Replace("\u2015", "-");
  text = text.Replace("\u2017", "_");
  text = text.Replace("\u2018", "'");
  text = text.Replace("\u2019", "'");
  text = text.Replace("\u201a", ",");
  text = text.Replace("\u201b", "'");
  text = text.Replace("\u201c", '"');
  text = text.Replace("\u201d", '"');
  text = text.Replace("\u201e", '"');
  text = text.Replace("\u2026", "...");
  text = text.Replace("\u2032", "'");
  text = text.Replace("\u2033", '"');

  text = text.Replace("\\", "<br/>");

  return text;
}

function GetFormatSize(value, decimalPlaces = 1) {
  let SizeSuffixes = ["bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"];

  if (value < 0) {
    return "-" + this.formatSize(-value);
  }

  let i = 0;
  let dValue = value;
  while (Math.Round(dValue, decimalPlaces) >= 1000) {
    dValue /= 1024;
    i++;
  }
  return `${dValue + decimalPlaces} ${SizeSuffixes[i]}`;
}
