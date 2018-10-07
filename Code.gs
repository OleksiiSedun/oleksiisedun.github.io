var TO_ADDRESS = "asedun@mackiev.com";

// function to spit out all the keys/values from the form in HTML
function formatMailBody(obj) { 
  var report = "";
  for (var key in obj) { 
    if (key == "QAC") {
      continue
    }
    if (key.indexOf("_D") == -1 && /\d/.test(obj[key])) {
      report += "<h4 style='text-transform: capitalize; margin-bottom: 0'>" + key + " "+ obj[key] + "</h4><div>" + obj[key + "_D"] + "</div>";
    }
  }
  return report; 
}

function doPost(e) {
  try {
    Logger.log(e); // the Google Script version of console.log see: Class Logger
    record_data(e);

    var mailData = e.parameters; // just create a slightly nicer variable name for the data
    var qaCode = mailData["QAC"]
    
    MailApp.sendEmail({
      to: TO_ADDRESS,
      subject: "QA Daily Report from " + qaCode,
      htmlBody: formatMailBody(mailData)
    });

    return ContentService   
          .createTextOutput(
            JSON.stringify({"result":"success",
                            "data": JSON.stringify(e.parameters)}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    Logger.log(error);
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  }
}

// record_data inserts the data received from the html form submission
// e is the data received from the POST
function record_data(e) {
  Logger.log(JSON.stringify(e)); // log the POST data in case we need to debug it
  try {
    var doc     = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = doc.getSheetByName('Reports'); 
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow() + 1;
    var row     = [new Date()]; // first element in the row should always be a timestamp
    // loop through the header columns
    for (var i = 1; i < headers.length; i++) { // start at 1 to avoid Timestamp column
      if(headers[i].length > 0) {
        row.push(e.parameter[headers[i]]); // add data to row
      }
    }
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
  }
  catch(error) {
    Logger.log(e);
  }
  finally {
    return;
  }
}
