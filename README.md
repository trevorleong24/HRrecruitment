function sendFeedbackEmails() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var 
 data = dataRange.getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var candidateName = row[0];
    var emailAddress = row[1];
    var feedbackMessage = row[2];
    var status = row[3];

    if (status == "Not Sent") {
      var message = GmailApp.createMessage()
                        .setTo(emailAddress)
                        .setSubject("Feedback on Your Application")
                        .setBody(feedbackMessage)
                        .send();
      sheet.getRange(i + 1, 4).setValue("Sent");
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Send Feedback Emails')
      .addItem('Send Emails', 'sendFeedbackEmails')
      .addToUi();
}

function sendFeedbackEmails() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var candidateName = row[0];
    var emailAddress = row[1];
    var feedbackMessage = row[2];
    var status = row[3];

    if (status == "Not Sent") {
      try {
        var message = GmailApp.createMessage()
                          .setTo(emailAddress)
                          .setSubject("Feedback on Your Application")
                          .setBody(feedbackMessage)
                          .send();
        sheet.getRange(i + 1, 4).setValue("Sent");
      } catch (error) {
        Logger.log('Error sending email to ' + emailAddress + ': ' + error.message);
        sheet.getRange(i + 1, 4).setValue("Error");
      }
    }
  }
}

function sendFeedbackEmails() {
  // ... (rest of the code)

  var emailTemplate = 'Dear {{candidateName}},\n\n{{feedbackMessage}}\n\nBest regards,\nHR Team';

  // ... (rest of the code)

  var message = GmailApp.createMessage()
                    .setTo(emailAddress)
                    .setSubject("Feedback on Your Application")
                    .setBody(emailTemplate.replace('{{candidateName}}', candidateName)
                                           .replace('{{feedbackMessage}}', feedbackMessage))
                    .send();
}

function sendFeedbackEmails() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  var emailTemplate = 'Dear {{candidateName}},\n\n{{feedbackMessage}}\n\nBest regards,\nHR Team';

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var candidateName = row[0];
    var emailAddress = row[1];
    var feedbackMessage = row[2];
    var status = row[3];

    if (status == "Not Sent") {
      try {
        var message = GmailApp.createMessage()
                          .setTo(emailAddress)
                          .setSubject("Feedback on Your Application")
                          .setBody(emailTemplate.replace('{{candidateName}}', candidateName)
                                           .replace('{{feedbackMessage}}', feedbackMessage))
                          .send();
        sheet.getRange(i + 1, 4).setValue("Sent");
      } catch (error) {
        Logger.log('Error sending email to ' + emailAddress + ': ' + error.message);
        sheet.getRange(i + 1, 4).setValue("Error");
      }
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Send Feedback Emails')
      .addItem('Send Emails', 'sendFeedbackEmails')
      .addToUi();
}

