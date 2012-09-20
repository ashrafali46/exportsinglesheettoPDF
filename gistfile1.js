// Simple function to send Weekly Status Sheets to contacts listed on the "Contacts" sheet in the MPD.

// Load a menu item called "Project Admin" with a submenu item called "Send Status"
// Running this, sends the currently open sheet, as a PDF attachment
function onOpen() {
  var submenu = [{name:"Send Status", functionName:"exportSomeSheets"}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Project Admin', submenu);  
}

function exportSomeSheets() {
  // Set the Active Spreadsheet so we don't forget
  var originalSpreadsheet = SpreadsheetApp.getActive();
  
  // Set the message to attach to the email.
  var message = "Please see attached"; // Could make it a pop-up perhaps, but out of wine today
  
  // Get Project Name from Cell A1
  var projectname = originalSpreadsheet.getRange("A1:A1").getValues(); 
  // Get Reporting Period from Cell B3
  var period = originalSpreadsheet.getRange("B3:B3").getValues(); 
  // Construct the Subject Line
  var subject = projectname + " - Weekly Status Sheet - " + period;

      
  // Get contact details from "Contacts" sheet and construct To: Header
  // Would be nice to include "Name" as well, to make contacts look prettier, one day.
  var contacts = originalSpreadsheet.getSheetByName("Contacts");
  var numRows = contacts.getLastRow();
  var emailTo = contacts.getRange(2, 2, numRows, 1).getValues();

  // Google scripts can't export just one Sheet from a Spreadsheet
  // So we have this disgusting hack

  // Create a new Spreadsheet and copy the current sheet into it.
  var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var projectname = SpreadsheetApp.getActiveSpreadsheet();
  sheet = originalSpreadsheet.getActiveSheet();
  sheet.copyTo(newSpreadsheet);
  
  // Find and delete the default "Sheet 1", after the copy to avoid triggering an apocalypse
  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();
  
  // Make zee PDF, currently called "Weekly status.pdf"
  // When I'm smart, filename will include a date and project name
  var pdf = DocsList.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
  var attach = {fileName:'Weekly Status.pdf',content:pdf, mimeType:'application/pdf'};

  // Send the freshly constructed email 
  MailApp.sendEmail(emailTo, subject, message, {attachments:[attach]});
  
  // Delete the wasted sheet we created, so our Drive stays tidy.
  DocsList.getFileById(newSpreadsheet.getId()).setTrashed(true);  
}