// This moves a row of data to the Tracker Archive sheet when "ARCHIVE" is written in uppercase in the Archive column.

function archive_rows(e){
  var row_no = e.range.getRow()-1;
  var row_data = SpreadsheetApp.getActiveSheet().getDataRange().getValues()[row_no];
  var archive_sheet=ss.getSheetByName('Tracker Archive');
  
  if(row_data[15]=="ARCHIVE"){  // Column 15 trigger
      Logger.log(row_data);
      archive_sheet.appendRow(row_data);
      ss.deleteRow(row_no+1);
    }
}

function createOnEditTrigger(e) {
  var triggers = ScriptApp.getProjectTriggers();
  var shouldCreateTrigger = true;
  triggers.forEach(function (trigger) {  
    
      if(trigger.getEventType() === ScriptApp.EventType.ON_EDIT && trigger.getHandlerFunction() === "archive_rows") {
      shouldCreateTrigger = false; 
    }
    
  });
 
  if(shouldCreateTrigger){
    ScriptApp.newTrigger("archive_rows").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  }
}
