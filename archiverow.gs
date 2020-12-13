// Moves row to "Tracker Archive" sheet and deletes when "Archive" is written/chosen from the dropdown in the first column of the row.

function onEdit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var r = ss.getActiveRange();
  var rows = r.getRow();
  var cell = s.getRange(rows, r.getColumn());

  if (s.getName() == "Tracker" && r.getColumn() == 1 && cell.getValue() == "Archive") { // "Tracker" is the sheet it will work out of, "1" is the column where it searches for the trigger to move the row, "Archive" is the value that it searches for to trigger the script to move the specified row.
  var numColumns = s.getLastColumn();
  var targetSheet = ss.getSheetByName("Tracker Archive"); // "Tracker Archive" is the row that it targets or moves the data to.
  var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
  s.getRange(rows, 1, 1, numColumns).moveTo(target);
  s.deleteRow(rows);
 }
}
