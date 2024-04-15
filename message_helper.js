
/**
 * This file contains Outlook Items related classes and related functions.
 * 
 */

const ERROR_TYPES = {
  api_error: "API_FAILED",
  ssl_error: "SSL_ERROR",
  generic_error: "GENERIC_ERROR"
}

/**
 * Email Attachment description
 */
class Attachment {

    constructor(name, content, contentFormat) {
      this.name = name;
      this.content = content;
      this.contentFormat = contentFormat
    }
  }
  
  /**
   * Message description
   */
class Message {
  
    constructor(messageType) {
      this.version = 1;
      this.messageType = messageType;
      this.subject = "";
      this.body = "";
      this.bodyTextFormat = "";
      this.sender = "";
      this.sentTime = new Date().getTime(); //milliseconds
      this.attachments = [];
      this.recipients = [];
      this.location = "";
    }
  
    addAttachment(name, content, contentFormat) {
      let attachmentObj = new Attachment(name, content, contentFormat);
      this.attachments.push(attachmentObj);
    }
  }


/**
 * Get details from current MailBox Item and 
 * generate the Message object
 */
class MessageGenerator
{
  constructor(item) {
  this.mailItem = item;
  }

  generateMessage(type)
  {
    messageObj = new Message(type);

    if (type === "message") {
      var promises = [
        this.executeCallback(this.setSubject),
        this.executeCallback(this.setSender),
        this.executeCallback(this.setBody),
        this.executeCallback(this.setTORecipients),
        this.executeCallback(this.setCCRecipients),
        this.executeCallback(this.setBCCRecipients),
        this.executeCallback(this.setAttachments) ];
      return Promise.all(promises);
    }
  }

  executeCallback(executor)
  {
    // using arrow function below to pass 'this' context to callback function
    return new Promise((resolve, reject) => {
      //executor(resolve, reject);
      executor.call(this, resolve, reject);  // this way we can pass 'this' context
      })  
  }

  setSubject(resolve, reject)
  {
    mailboxItem.subject.getAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        messageObj.subject = asyncResult.value; 
        resolve();    
      } else {
          reject(asyncResult.error);
        }       
    });
  }

  setSender(resolve, reject)
  {
    mailboxItem.from.getAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        var emailAddressDetail = asyncResult.value;
        messageObj.sender = emailAddressDetail.emailAddress;
        resolve();
      } else {
          reject(asyncResult.error);
      }
    });
  }

  async setBody(resolve, reject)
  {
    // For outlook client, using message body as 'text'
    let bodyType = Office.CoercionType.Text;
    if (officeHostName === OFFICE_HOST_NAMES.OUTLOOK_WEB_ACCESS){
        try{
          bodyType = await this.getBodyTypeAsync();
        } catch (error){
            reject(error);
      }
    }
    messageObj.bodyTextFormat = bodyType;
    //console.log("Body type recieved : " + bodyType);

    mailboxItem.body.getAsync(bodyType, function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        messageObj.body = asyncResult.value;
        resolve();
      } else {
          reject(asyncResult.error);
      }
    });
  }

  getBodyTypeAsync(){
    return new Promise(function (resolve, reject) {
      mailboxItem.body.getTypeAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(asyncResult.value);
        } else {
            reject(asyncResult.error);
        }
      });
    });
  }

  setTORecipients(resolve, reject)
  {
    mailboxItem.to.getAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) { 
        var emailAddressDetails = asyncResult.value;

        emailAddressDetails.forEach(function (emailAddressDetail){
          messageObj.recipients.push(emailAddressDetail.emailAddress);
        })
        resolve();  
      } else {
         reject(asyncResult.error);
      }
    });
  }

  setCCRecipients(resolve, reject)
  {
    mailboxItem.cc.getAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) { 
        var emailAddressDetails = asyncResult.value;

        emailAddressDetails.forEach(function (emailAddressDetail){
          messageObj.recipients.push(emailAddressDetail.emailAddress);
        })
        resolve();  
      } else {
          reject(asyncResult.error);
      }
    });
  }

  setBCCRecipients(resolve, reject)
  {
    mailboxItem.bcc.getAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) { 
        var emailAddressDetails = asyncResult.value;

        emailAddressDetails.forEach(function ( emailAddressDetail ){
          messageObj.recipients.push( emailAddressDetail.emailAddress );
        })
        resolve();  
      } else {
          reject(asyncResult.error);
      }
    });
  }

  setLocation(resolve, reject)
  {
    mailboxItem.location.getAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) { 
        messageObj.location = asyncResult.value;
        resolve();
      } else {
          reject(asyncResult.error);
      }
    });
  }


  setAttachments(resolve, reject)
  {
	console.log("In setAttachments: " + new Date().toString());  
    mailboxItem.getAttachmentsAsync((asyncResult) => {  // using arrow function to pass 'this' context
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) { 
      let attachmentDetailsList = asyncResult.value; //this returns an array of type Office.AttachmentDetails
      
      if (attachmentDetailsList.length > 0) {
        
        let attachmentPromises = [];
        for (let i = 0; i < attachmentDetailsList.length ; i++) {
          var attachmentDetails = attachmentDetailsList[i];
          if(attachmentDetails.attachmentType === "cloud") {
               continue;
          }
          attachmentPromises.push(this.processAttachment(attachmentDetails));
        }

        // We need to get content of all the attachments, so using Promise.all()
        Promise.all(attachmentPromises).then(() =>{
          resolve(); 
        });

      } else {
		  console.log("setAttachments: attachmentPromises before resolve: " + new Date().toString());
          resolve();
        }
    } else {
        reject(asyncResult.error);
      }
    });
  }  

  processAttachment(attachmentDetails){
    return new Promise(function (resolve, reject) {
	  console.log("processAttachment: Calling getAttachmentContentAsync  " + new Date().toString());
      mailboxItem.getAttachmentContentAsync(attachmentDetails.id, (getAttachmentContentResult) =>{
      console.log("processAttachment: Inside getAttachmentContentAsync  " + new Date().toString());
        let attachmentContent = getAttachmentContentResult.value; //Office.AttachmentContent interface
        let content = "";
  
        switch(attachmentContent.format) {
          case Office.MailboxEnums.AttachmentContentFormat.Base64:
            content = attachmentContent.content;
            break;
          case Office.MailboxEnums.AttachmentContentFormat.Eml:
            content = attachmentContent.content;
            break;
          case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
            content = attachmentContent.content;
            break;
          case Office.MailboxEnums.AttachmentContentFormat.Url:
            // Nothing to do here
            break; 
          default:
            // Handle attachment formats that are not supported.
        }
        if(content.length !== 0) {
		  console.log("processAttachment: Before messageObj.addAttachment  " + new Date().toString());
          messageObj.addAttachment(attachmentDetails.name, content, attachmentContent.format);
		  console.log("processAttachment: After messageObj.addAttachment  " + new Date().toString());
        }
		console.log("processAttachment: Before resolve processAttachment Promise  " + new Date().toString());
        resolve();
    });
  });
  }

}

class AddInError {

  constructor(type, name, description) {
    this.type = type;
    this.name = name;
    this.description = description;
    
  }
}

class ErrorMessage {

  constructor() {
    this.errors = [];
    this.eventtime = new Date().getTime(); //milliseconds
  }

  addError(addInError){
    this.errors.push(addInError);
  }

}
