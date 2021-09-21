var EMAIL_COL = 1;
var EMAIL_SENT_COL = 5;
var EMAIL_SENT = 'EMAIL_SENT';
var DEFAULT_EMAIL = 1;
var SURVEY_TIME = 4;
var INITIAL_DATE_COL = 5;
var FINAL_DATE_COL = 19;


function checkEmailStatus()
{
  const scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  //remove the top row and column headers
  const scheduleData = scheduleSheet.getDataRange().getValues().slice(2);
  scheduleData.forEach(x => {
    const emailAddress = x[EMAIL_COL];
    const surveyTime = x[SURVEY_TIME];
    const currentTime = new Date();
    let nextEmail = undefined;
    for(let i = INITIAL_DATE_COL; i <= FINAL_DATE_COL; i += 2)
    {
      //check if the 'email successful' column been set 
      if(!x[i+1])
      {
        //nextEmail = new Date(x[i]).setHours(surveyTime);
        nextEmail = new Date(x[i]);
        break;
      }
    }
    if(currentTime > nextEmail){
       console.log(emailAddress+' '+nextEmail);
      //sendEmail(emailAddress)
    }
  });
}

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
  var initialDate = new Date(row[INITIAL_DATE_COL]);
  for(let i = 1; i < 8; i++)
  {
    const nextDate = new Date();
    nextDate.setDate(initialDate.getDate() + i*7);
    console.log(nextDate);
    dataSheet.getRange(range.getRow(), INITIAL_DATE_COL + i*2 + 1).setValue(Utilities.formatDate(nextDate, "EST", "MM/dd/yyyy"));
  }
}
