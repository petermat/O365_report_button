/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  
  // =================================================
  // Change email address HERE
  var toAddress = "email-analysis@eset.nl";
  // =================================================

  const message_done = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Email Reported to Security Team.",
    icon: "Icon.80x80",
    persistent: true
  };

  const message_manual = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Email Forwarded to Security Team.",
    icon: "Icon.80x80",
    persistent: true
  };

  const message_failed = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Email Submission Failed.",
    icon: "Icon.80x80",
    persistent: true
  };

  function reportMessage() {
    var item = Office.context.mailbox.item;
    itemId = item.itemId;
    mailbox = Office.context.mailbox;
    try{
        easyEws.getMailItemMimeContent(itemId, sendMessageCallback, showErrorCallback);
    } catch (error) {
      console.log("request failed");
    }
}

function sendMessageCallback(content) {
  try{
      easyEws.sendPlainTextEmailWithAttachment("[REPORTED EMAIL] \"" + Office.context.mailbox.item.normalizedSubject + "\"",
      "A user has forwarded a suspicious email",
                                               toAddress,
                                               "DANGER_reported_email.eml",
                                               content,
                                               successCallback,
                                               showErrorCallback);
  }
  catch (error) {
    console.log(error.message);
  }
}

// This function is the callback for the easyEws sendPlainTextEmailWithAttachment
// Recieves: a message that the result was successful.
function successCallback(result) {
  //showNotification("Success", result);
    // Show a notification message
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", message_done);

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

// This function will display errors that occur 
// we use this as a callback for errors in easyEws
function showErrorCallback(error) {
  //console.log("Failback to manual forwarding");
  var parentEmail = Office.context.mailbox.item;

  // Office.context.mailbox.userProfile.emailAddress
  Office.context.mailbox.displayNewMessageFormAsync(
  {
        toRecipients: [toAddress,], // Office.context.mailbox.item.to Copies the To line from current item
        subject: "[REPORTED EMAIL] \"" + Office.context.mailbox.item.normalizedSubject + "\"",
        htmlBody : "<p>To report phishing emails you've received, please forward this email for analysis with <strong>Send</strong> button below.</p><br><p>Please do not remove the attachment.</p>",
        attachments :
        [
            { type: "item", itemId : Office.context.mailbox.item.itemId, name: "DANGER_reported_email.eml", isInline: false }
        ]},
        function(asyncResult){
          //console.log(JSON.stringify(asyncResult));
          if (asyncResult.status == "succeeded")
          {
            parentEmail.notificationMessages.replaceAsync("action", message_manual);
            event.completed();
           } else  {
            parentEmail.notificationMessages.replaceAsync("action", message_failed); 
            //console.log("Action failed with error: " + asyncResult.error.message);
           }
        }        
    );
}
reportMessage();

}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
