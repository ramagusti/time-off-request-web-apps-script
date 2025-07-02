var globalIdSpreadsheet = "xxx"
var globalId
var urlGForm = "https://docs.google.com/forms/d/e/xxx/viewform?usp=sf_link"

function doGet(e) {
  globalId = e.parameter['id'];

  const html = HtmlService.createTemplateFromFile(e.parameter['status']).evaluate().addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  const emailRequester = e.parameter['email'];
  const status = e.parameter['status'].toString().replaceAll("-", " ");
  const prevStatus = [].concat.apply([], SpreadsheetApp.openById(globalIdSpreadsheet).getSheetByName("Time Off Requests").getRange("H2:I").getValues().filter(val=>val[1] == globalId))[0];
  const calendarSheet = SpreadsheetApp.openById(globalIdSpreadsheet).getSheetByName("Calendar IDs");
  
  if (prevStatus != status.toProperCase()) {
    setStatus(globalId, status.toProperCase());
    handleApproval(globalId,status,emailRequester)

    const deleted = calendarSheet.getRange("A2:B").getValues().filter(val=>val[0] && val[0] == globalId).map(val=>val[1]);
    console.log(deleted);
    if (deleted.length > 0) {
      for (id of deleted) {
        CalendarApp.getEventById(id).deleteEvent();
      }
      const cleaned = calendarSheet.getRange("A2:B").getValues().filter(val=>val[0] && val[0] != globalId).map(val=>[val[0], val[1]]);
      calendarSheet.getRange("A2:B").clearContent();
      if (cleaned.length > 0) {
        calendarSheet.getRange(2, 1, cleaned.length, 2).setValues(cleaned);
      }
    }
  }
  
  if (status == "approved") {
    if (calendarSheet.getRange("A2:A").getValues().filter(val=>val[0] == globalId).length == 0) {
      setCalendar(globalId);
    }
  }
  
  return html.setTitle('Confirmation');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  var url = ScriptApp.getService().getUrl();
  console.log("Request URL: " + url);
  return url;
}

function setStatus(id, status) {
  
  const sheet = SpreadsheetApp.openById(globalIdSpreadsheet).getSheetByName("Time Off Requests");
  const row = sheet.getRange(`I2:I${sheet.getLastRow()}`).getValues().map(val=>val[0]).indexOf(id) + 2;

  sheet.getRange(row, 8).setValue(status);
  SpreadsheetApp.flush();
}



function setCalendar(id) {
  const sheet = SpreadsheetApp.openById(globalIdSpreadsheet).getSheetByName("Time Off Requests");
  const row = sheet.getRange(`I2:I${sheet.getLastRow()}`).getValues().map(val=>val[0]).indexOf(id) + 2;
  const name = sheet.getRange(row, 3).getValue();
  const startDate = new Date(sheet.getRange(row, 4).getValue());
  const endDate = new Date(sheet.getRange(row, 5).getValue());
  const type = sheet.getRange(row, 6).getValue();

  const calendar = CalendarApp.createAllDayEvent(`${name} ${type}`, startDate, endDate);
  const calendarSheet = SpreadsheetApp.openById(globalIdSpreadsheet).getSheetByName("Calendar IDs");
  calendarSheet.getRange(getLastRow_(calendarSheet, 1) + 1, 1, 1, 2).setValues([[id, calendar.getId()]]);
}

function handleApproval(id,status,requester)
{
  globalId = id
  console.log(status);
  var template
    
  if (status == "approved") {
    template = "request-approve"
  }else if (status == "denied"){
    template = "request-denied"
  }else if (status == "requested more information"){
    template = "request-moreinfo"
  }

  var templ = HtmlService.createTemplateFromFile(template);
  var message = templ.evaluate().getContent();

  MailApp.sendEmail({
    to: requester,
    subject: "Time Off Request!",
    htmlBody: message
  });
    
}

function getDataForm() {

    const sheet = SpreadsheetApp.openById(globalIdSpreadsheet).getSheetByName("Time Off Requests");
    const row = sheet.getRange(`I2:I${sheet.getLastRow()}`).getValues().map(val=>val[0]).indexOf(globalId) + 2;


  let data = {
    appliedDate: Utilities.formatDate(sheet.getRange(row, 1).getValue(), 'America/Los_Angeles', 'MMMM dd, yyyy'),
    name: sheet.getRange(row, 3).getValue(),
    email: sheet.getRange(row, 2).getValue(),
    startingOn: Utilities.formatDate(sheet.getRange(row, 4).getValue(), 'America/Los_Angeles', 'MMMM dd, yyyy'),
    endOn:  Utilities.formatDate(sheet.getRange(row, 5).getValue(), 'America/Los_Angeles', 'MMMM dd, yyyy'),
    typeOfLeave: sheet.getRange(row, 6).getValue(),
    reasonLeave: sheet.getRange(row, 7).getValue(),
    urlgform: urlGForm,
};

  SpreadsheetApp.flush();

  return data;
}

String.prototype.toProperCase = function () {
  return this.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
};

function getLastRow_(sheet, columnNumber) {
  // version 1.5, written by --Hyde, 4 April 2021
  const values = (
    columnNumber
      ? sheet.getRange(1, columnNumber, sheet.getLastRow() || 1, 1)
      : sheet.getDataRange()
  ).getDisplayValues();
  let row = values.length - 1;
  while (row && !values[row].join('')) row--;
  return row + 1;
}
