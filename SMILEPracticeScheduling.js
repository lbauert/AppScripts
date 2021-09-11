var EMAIL_COL = 1;
var EMAIL_SENT_COL = 5;
var EMAIL_SENT = 'EMAIL_SENT';
var DEFAULT_EMAIL = 1;

var INITIAL_DATE = 7;
/*
* Sends email, given the range that
* triggered the email and they email type.
*/
function sendEmail(email) {
  try {
    // get data from range
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var range = ss.getActiveRange();
    var dataSheet = ss.getSheets()[0];
    var dataRange = dataSheet.getRange(range.getRow(), 1, 1, range.getColumn() + 15);
    var row = dataRange.getValues()[0];
    // get relevent columns
    var emailAddress = row[EMAIL_COL];
    var emailSent = row[EMAIL_SENT_COL - 1]; // assumes verification column always follows activation column
    // get email template
    var templateSheet = ss.getSheets()[1];
    var emailTemplateSheet = templateSheet.getRange(DEFAULT_EMAIL, 1, 1, 3);
    var defaultTemplate = emailTemplateSheet.getValues()[0];
    // check if already sent
    if (emailSent != EMAIL_SENT) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: emailTemplateSheet[0],
        htmlBody: emailTemplateSheet[2],
        //inlineImages: picArgs
      });
      dataSheet.getRange(range.getRow(), EMAIL_SENT_COL).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
      range.setNote('Activated: ' + new Date()); 
      return true;
    } else {
      displayError("This participant has already recieved this type of email. No email sent.");
      return false;
    }
  } catch (e) {
    displayError(e);
  }
}

function populateSurveySchedule()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getActiveRange();
  var dataSheet = ss.getSheets()[0];
  var dataRange = dataSheet.getRange(range.getRow(), 1, 1, range.getColumn() + 15);
  var row = dataRange.getValues()[0];
  var initialDate = new Date(row[INITIAL_DATE]);
  for(let i = 1; i < 8; i++)
  {
    const nextDate = new Date();
    nextDate.setDate(initialDate.getDate() + i*7);
    dataSheet.getRange(range.getRow(), INITIAL_DATE + i*2 + 1).setValue(Utilities.formatDate(nextDate, "EST", "MM/dd/yyyy"));
  }
}
