// Archive a row of data to another sheet based on a trigger word and the placement of this word within a specific column (#2). Triggers can be specified in the var archiveStates

function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = e.range;
  var s = r.getSheet();

  var row = r.getRow();
  var col = r.getColumn();
  var cell = s.getRange(row, col);
  var archiveStates = ["Archive", "Cancel", "Canceled", "Cancelled"];

  if (s.getName() === "Tracker" && r.getColumn() === 2 && archiveStates.includes(cell.getValue())) {
    var numColumns = s.getLastColumn();
    var targetSheet = ss.getSheetByName("Tracker Archive"); // "Tracker Archive" is the 2nd/target sheet
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    s.getRange(row, 1, 1, numColumns).moveTo(target);
    s.deleteRow(row);
  }
  if (s.getName() === "Offers" && r.getColumn() === 2 && archiveStates.includes(cell.getValue())) {
    var numColumns = s.getLastColumn();
    var targetSheet = ss.getSheetByName("Offers Archive"); // "Offers Archive" is the 2nd/target sheet
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    s.getRange(row, 1, 1, numColumns).moveTo(target);
    s.deleteRow(row);
  }
}
