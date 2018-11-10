function gestion_erreur(e, message) {

  var errorSheet = SpreadsheetApp.openById('1NfD2NqvXn_rCBPGuwz2f0zmpl0xqY3kIOeuYhBdat10').getSheetByName('Erreurs');
  var lastRow = errorSheet.getLastRow();
  var cell = errorSheet.getRange('A1');
  
  var current_console = SpreadsheetApp.getActive();
  if(current_console != undefined) var current_console_name = current_console.getName();
  else var current_console_name = "unkown";
 
  cell.offset(lastRow, 0).setValue(Date());
  cell.offset(lastRow, 1).setValue(e.message);
  cell.offset(lastRow, 2).setValue(e.fileName);
  cell.offset(lastRow, 3).setValue(e.lineNumber);
  cell.offset(lastRow, 4).setValue(message);
  cell.offset(lastRow, 5).setValue(current_console_name);
  
   MailApp.sendEmail("vesela@indemniflight.com", "Error report", 
      "\r\nMessage: " + e.message
      + "\r\nFile: " + e.fileName
      + "\r\nLine: " + e.lineNumber
      + "\r\nLine: " + message);
}