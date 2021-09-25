var EMAIL_COL = 1;
var EMAIL_SENT_COL = 5;
var EMAIL_SENT = 'EMAIL_SENT';
var DEFAULT_EMAIL = 1;
var SURVEY_TIME = 4;
var INITIAL_DATE_COL = 5;
var FINAL_DATE_COL = 19;
var NUMBER_OF_HEADER_ROWS = 2;

function checkEmailStatus()
{
  const scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  //remove the top row and column headers

  const scheduleData = scheduleSheet.getDataRange().getValues().slice(NUMBER_OF_HEADER_ROWS);
  scheduleData.forEach((x, subjectIndex) => {
    const emailAddress = x[EMAIL_COL];
    const surveyTimeHours = Utilities.formatDate(x[SURVEY_TIME], "EST", "HH");
    const surveyTimeMinutes = Utilities.formatDate(x[SURVEY_TIME], "EST", "mm");
    let nextEmail;
    for(let i = INITIAL_DATE_COL; i <= FINAL_DATE_COL; i += 2)
    {
      //check if the 'email successful' column been set 
      if(!x[i+1])
      {
        //build full date time
        const emailDate = new Date(x[i])
        emailDate.setHours(surveyTimeHours);
        emailDate.setMinutes(surveyTimeMinutes);
        nextEmail = {
          date: emailDate,
          logFn: (message) => {
            const row = subjectIndex+NUMBER_OF_HEADER_ROWS;
            const column = i+1;
            //increment rows and columns because sheet isn't 0-indexed :[
            scheduleSheet.getRange(row + 1, column + 1).setValue(message);
          },
        };
        break;
      }
    }
    
    //only send email if its passed their next date + time
    if(new Date() > nextEmail.date){
      sendEmail(emailAddress, nextEmail.logFn);
    }
  });
}

/*
* Sends email, given the range that
* triggered the email and they email type.
*/
function sendEmail(email, logFn) {
  try {
    // get data from range
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // get email template
    var templateSheet = ss.getSheets()[1];
    var emailTemplateSheet = templateSheet.getRange(DEFAULT_EMAIL, 1, 1, 3);
    var defaultTemplate = emailTemplateSheet.getValues()[0];
    // check if already sent
    MailApp.sendEmail({
      to: email,
      subject: defaultTemplate[0],
      htmlBody: defaultTemplate[2],
    });
    logFn(EMAIL_SENT);
    // Make sure the cell is updated right away in case the script is interrupted
    SpreadsheetApp.flush();
    return true;
  } catch (e) {
    displayError(e);
    logFn(EMAIL_ERROR);
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
