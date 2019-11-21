function sendEmail() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = sheet.getLastRow()-1;   // Number of rows to process
  // method getRange(row, column, optNumRows, optNumColumns)
  var dataRange = sheet.getRange(startRow, 1, numRows,3);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  var update_cell = sheet.getRange('H5');
  var subject ='';
  var message ='';

  for (i in data) {
    var row = data[i];
    var date = new Date();
    var sheetDate = new Date(row[1]);
    Sdate=Utilities.formatDate(date,'GMT-0500','yyyy-MM-dd')
    SsheetDate=Utilities.formatDate(sheetDate,'GMT-0500', 'yyyy-MM-dd')
    Logger.log('dates: ' + Sdate+' =? '+SsheetDate)

    if (Sdate == SsheetDate){
      var emailAddress = SpreadsheetApp.getActiveSheet().getRange('H2').getValue();     // Get Email to send to
       subject = SpreadsheetApp.getActiveSheet().getRange('H3').getValue() + row[0];    // Get default subject

      // Check if custom message exists
      if ( row[2] =='')
         {
             message = SpreadsheetApp.getActiveSheet().getRange('H4').getValue();
          }
      else {
              message =row[2];
            }

      MailApp.sendEmail(emailAddress, subject, message);
      Logger.log('SENT: '+emailAddress+'  '+subject+'  '+message)
      // Quick audit
      update_cell.setValue(SsheetDate);
    }
  }
}
