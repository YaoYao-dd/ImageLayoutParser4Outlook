/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

let serverUrl = "https://localhost:8000/uploadImgs/"

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});


function getEmailImageUrls(emailContent) {
    var parser = new DOMParser();
    var doc = parser.parseFromString(emailContent, "text/html");
    var imageElements = doc.getElementsByTagName("img");
    return imageElements;
}

async function getImages() {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var emailContent = result.value;
      var imgElements = getEmailImageUrls(emailContent);
      for(var i=0; i<imgElements.length; i++){
        var imgElement = imgElements[i];
        var imageUrl = imgElement.src;
        if(imageUrl.indexOf("cid") > -1){
          fetchAttachments(imageUrl.split(":")[1], imgElement.width, imgElement.height);
        }
        else{
          
        }
      }
    } else {
      console.error(result.error);
    }
  });
}

function getItemRestId(itemId) {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
      return itemId;
  } else {
    return Office.context.mailbox.convertToRestId(
      itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}

async function fetchAttachments(contentId, width, height) {
  Office.context.mailbox.getCallbackTokenAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var accessToken = result.value;
      var messageId = Office.context.mailbox.item.itemId;
      var itemId = getItemRestId(messageId);
      var restUrl = Office.context.mailbox.restUrl;
      var attachmentsUrl = restUrl + '/v2.0/me/messages/' + itemId + '/attachments';
      console.log(attachmentsUrl);
      console.log(attachmentsUrl);
      $.ajax({
        url: attachmentsUrl,
        type: 'GET',
        headers: {
          'Authorization': 'Bearer ' + accessToken
        },
        success: function(response) {
          var attachments = response.value;
          for (var i = 0; i < attachments.length; i++) {
            var attachment = attachments[i];
            console.log(attachment);
            if(attachment.ContentId != contentId)
            {
              continue;
            }
            $.ajax({
              url: serverUrl,
              method: 'POST',
              contentType: 'application/json',
              data: JSON.stringify({}),
              success: function(response) {
                var img = base64ToImage(attachment.ContentBytes);
                var mapName = "Test";
                img.setAttribute("usemap", "#" + mapName);
                img.setAttribute("alt", "Outter help text");
                var mapElement = initMap(response, mapName);
                document.getElementById("content").append(img);
                document.getElementById("content").append(mapElement);
                // console.log(response);
                // console.log(img);
                // console.log(width + ":" +img.width);
                // console.log(height + ":" + img.height);
                // let utterance = new SpeechSynthesisUtterance("Hello world!");
                // speechSynthesis.speak(utterance);
                showModalDialog();
              }
            });
            
          }
        },
        error: function(error) {
          console.error(error);
        }
      });
    } else {
      console.error(result.error);
    }
  });
}

function base64ToImage(base64String) {
  const binaryString = atob(base64String);
  const img = document.createElement('img');
  img.setAttribute("style","max-width: 100%;height: auto;");
  img.src = `data:image/png;base64,${btoa(binaryString)}`;
  return img;
}

function initMap(response, mapName){
  const mapElement = document.createElement('map');
  mapElement.setAttribute("id", mapName);
  mapElement.setAttribute("name", mapName);
  var segments = response.segments;
  for(var i = 0 ; i < segments.length; i++){
    var segment = segments[i];
    const area = document.createElement('area');
    area.setAttribute("shape", "rect");
    area.setAttribute("coords", segment["coordinates"]);
    area.setAttribute("alt", segment["text"]);
    area.setAttribute("href", "#");
    mapElement.appendChild(area);
  }
  return mapElement;
}

var dialog;

function showModalDialog() {
  var dialogUrl = window.location.origin + '/taskpane.html';
  var dialogOptions = { height: 300, width: 400, displayInIframe: true };

  dialog = Office.context.ui.displayDialogAsync(dialogUrl, dialogOptions, function(result) {
    dialog = result.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(args) {
      console.log(args.message);
      // Process the message received from the dialog
    });
  });
}

export async function run() {
  var contentEle = document.getElementById("content");
  while(contentEle.hasChildNodes()) {
    contentEle.removeChild(contentEle.firstChild);
  }
  getImages();
}
