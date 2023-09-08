/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});


function getEmailImageUrls(emailContent) {
    var parser = new DOMParser();
    var doc = parser.parseFromString(emailContent, "text/html");
    var imageElements = doc.getElementsByTagName("img");
    var imageUrls = [];

    for (var i = 0; i < imageElements.length; i++) {
      var imageUrl = imageElements[i].src;
      imageUrls.push(imageUrl);
    }
    return imageUrls;
}

function getImages() {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var emailContent = result.value;
      var imageUrls = getEmailImageUrls(emailContent);
      // Use the image URLs here
      console.log(imageUrls);
    } else {
      // Handle the error
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

function fetchAttachments() {
  Office.context.mailbox.getCallbackTokenAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var accessToken = result.value;
      var messageId = Office.context.mailbox.item.itemId;
      var itemId = getItemRestId(messageId);
      var restUrl = Office.context.mailbox.restUrl;
      var attachmentsUrl = restUrl + '/v2.0/me/messages/' + itemId + '/attachments';
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
            console.log(attachment.ContentBytes);
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

function convertRemoteImageToBase64(imageUrl, callback) {
  fetch(imageUrl)
    .then(response => response.blob())
    .then(blob => {
      var reader = new FileReader();
      reader.onloadend = function() {
        var base64String = reader.result.split(',')[1];
        callback(base64String);
      };
      reader.readAsDataURL(blob);
    })
    .catch(error => {
      console.error(error);
    });
}


export async function run() {
  fetchAttachments();
}
