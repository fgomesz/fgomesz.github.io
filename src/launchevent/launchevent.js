/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

function onNewMessageComposeHandler(event) {
  setSubject(event);
}
function onNewAppointmentComposeHandler(event) {
  setSubject(event);
}
function onMessageComposeHandler(event) {
  sayHello(event);
}
function setSubject(event) {
  Office.context.mailbox.item.subject.setAsync(
    "Set by an event-based add-in!",
    {
      asyncContext: event,
    },
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
      }

      // Call event.completed() after all work is done.
      asyncResult.asyncContext.completed();
    }
  );
}
function sayHello() {
  // Get the current context
  var context = Office.context;
  console.log("whatgoing on.mailbox.item: ", context.mailbox.item);

  if (context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
    Office.context.mailbox.item.to.getAsync(function (asyncResult) {
      if (asyncResult.status === "succeeded") {
        var recipients = asyncResult.value;
        var greeting = "Hello"
        
        if (recipients.length > 0) {
          
          //for (var i = 0; i < recipients.length; i++) {
            var mailName = recipients[i].displayName;
            
            try {
              var mailSplit = mailName;
              if (mailName.includes(",")) {
                mailSplit = mailName.split(",");
              } else if (mailName.includes("\/QMPD")) {
                greeting += " SRMD team";
                break; 
              } else {
                mailSplit = mailName.split(" ");
              }

              var lastName = mailSplit[0];
              var firstName = mailSplit[1];

              if ( (mailSplit === undefined) || (mailSplit == 'undefined')) {
                greeting += " " + mailName + ",";
              } else if (mailName.includes("(SEC)")) {
                greeting += " " + lastName + "-san,";
              } else if (mailName.includes("(Contractor)")) {
                firstName = firstName.replace("(Contractor)", "");
                greeting += " " + firstName + ",";
              } else {
                greeting += " " + firstName + ",";
              }
            } catch (error) {
              
              greeting += " " + recipients[i].displayName + "";
            }
          //} # for loop
          greeting += "\n\n";
          
          // Insert the greeting into the reply message body
          Office.context.mailbox.item.body.setSelectedDataAsync(
            greeting,
            {
              coercionType: Office.CoercionType.Text,
              asyncContext: {},
            },
            function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
              }
            }
          );
        } else {
          console.error("No recipients found");
        }
      } else {
        console.error("Failed to retrieve recipients: " + asyncResult.error.message);
      }
    });
  }
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
  Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
}
