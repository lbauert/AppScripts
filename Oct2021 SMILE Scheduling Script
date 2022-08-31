var EMAIL_COL = 1;
var EMAIL_SENT_COL = 5;
var EMAIL_SENT = 'EMAIL_SENT';
var SURVEY_TIME = 4;
var INITIAL_DATE_COL = 5;
var MID_POINT_DATE_COL = 11;
var FINAL_DATE_COL = 19;
var NUMBER_OF_HEADER_ROWS = 2;

function checkEmailStatus()
{
  const scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  //remove the top row and column headers
  const scheduleData = scheduleSheet.getDataRange().getValues().slice(NUMBER_OF_HEADER_ROWS);
  scheduleData.forEach((row, rowIndex) => {
    const emailAddress = row[EMAIL_COL];
    const surveyTimeHours = Utilities.formatDate(row[SURVEY_TIME], "EST", "HH");
    const surveyTimeMinutes = Utilities.formatDate(row[SURVEY_TIME], "EST", "mm");
    let nextEmail;
    for(let column = INITIAL_DATE_COL; column <= FINAL_DATE_COL; column += 2)
    {
      //check if the 'email successful' column been set 
      if(!row[column+1])
      {
        //build full date time
        const emailDate = new Date(row[column])
        emailDate.setHours(surveyTimeHours);
        emailDate.setMinutes(surveyTimeMinutes);
        nextEmail = {
          date: emailDate,
          logFn: (message) => {
            const logRow = rowIndex+NUMBER_OF_HEADER_ROWS;
            const logColumn = column+1;
            //increment rows and columns because sheet isn't 0-indexed :[
            scheduleSheet.getRange(logRow + 1, logColumn + 1).setValue(message);
          },
          emailInfo: getEmailTemplateInformation(column)
          
        };
        break;
      }
    }
    
    //only send email if its passed their next date + time
    if(new Date() > nextEmail.date){
      sendEmail(emailAddress, nextEmail.logFn, nextEmail.emailInfo);
    }
  });
}

function getEmailTemplateInformation(column)
{
  // get data from range
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheets()[1];
  var emailTemplateSheet = templateSheet.getRange(1, 1, 2, 3);
  let emailInfo;
  switch(column)
  {
    case MID_POINT_DATE_COL:
      emailInfo = emailTemplateSheet.getValues()[1];
      break;
    case FINAL_DATE_COL:
      emailInfo = emailTemplateSheet.getValues()[2];
      break;
    default:
      emailInfo = emailTemplateSheet.getValues()[0];
  }
  return emailInfo;
}


/*
* Sends email, given the range that
* triggered the email and they email type.
*/
function sendEmail(emailAddress, logFn, emailInfo) {
  try {     
    // check if already sent
    MailApp.sendEmail({
      to: emailAddress,
      subject: emailInfo[0],
      htmlBody: emailInfo[2],
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
  const scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  //remove the top row and column headers
  const scheduleData = scheduleSheet.getDataRange().getValues().slice(NUMBER_OF_HEADER_ROWS);
  scheduleData.forEach((row, rowIndex) => 
  {
    var initialDate = new Date(row[INITIAL_DATE_COL]);
    for(let i = 1; i < 8; i++)
    {
      if(!scheduleSheet.getRange(rowIndex+NUMBER_OF_HEADER_ROWS + 1, INITIAL_DATE_COL + i*2 + 1).getValue()){
        const nextDate = new Date(initialDate);
        nextDate.setDate(nextDate.getDate() + i*7);
        scheduleSheet.getRange(rowIndex+NUMBER_OF_HEADER_ROWS + 1, INITIAL_DATE_COL + i*2 + 1).setValue(Utilities.formatDate(nextDate, "EST", "MM/dd/yyyy"));
      }
    }
  });
}
