
const OFFICE_HOST_NAMES = {
  OUTLOOK_CLIENT: "Outlook",
  OUTLOOK_WEB_ACCESS: "OutlookWebApp"
}

let officeHostName;
let mailboxItem;
let messageObj;

Office.initialize = function (reason) {
  mailboxItem = Office.context.mailbox.item;
  officeHostName = Office.context.mailbox.diagnostics.hostName;
};

async function validateBody(event){

  var messageGeneratorObj = new MessageGenerator(mailboxItem);
  try {
    await messageGeneratorObj.generateMessage(mailboxItem.itemType);
  }
  catch (error){
    console.error(error);
  }

  var jsonMessage = JSON.stringify(messageObj);
  console.log("Json message : ");
  console.log(jsonMessage);

}

async function main(event) {
  try {
	console.log("Inside Main");
    await validateBody(event);
	console.log("After Validate Body Exiting.");
	event.completed({allowEvent: true});
  } catch (error) {
    console.error("Error occurred while processing email message. Allowing Send event.");
    console.error(error);
    event.completed({allowEvent: true});
  }
}

