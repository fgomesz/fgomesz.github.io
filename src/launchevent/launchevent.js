function sayHello() {
  // Get the current context
  var context = Office.context;
  console.log("mailbox.item: ", context.mailbox.item);

  if (context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
    Office.context.mailbox.item.to.getAsync(function (asyncResult) {
      if (asyncResult.status === "succeeded") {
        var recipients = asyncResult.value;
        var greeting = "Hello";
        
        if (recipients.length > 0) {
          var mailName = recipients[0].displayName; // Only handle the first recipient

          try {
            var mailSplit;
            if (mailName.includes(",")) {
              mailSplit = mailName.split(",");
            } else if (mailName.includes("/QMPD")) {
              greeting += " SRMD team";
            } else {
              mailSplit = mailName.split(" ");
            }

            var lastName = mailSplit ? mailSplit[0] : "";
            var firstName = mailSplit ? mailSplit[1] || "" : "";

            if (!mailSplit || mailSplit === 'undefined') {
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
            greeting += " " + mailName;
          }

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
