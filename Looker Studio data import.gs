function copyDataToSheet2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Calendars data extraction tool");
  var targetSheet = ss.getSheetByName("Looker data");

  // Get data range from source sheet
  var sourceRange = sourceSheet.getRange("A2:H");

  // Get values from the range
  var data = sourceRange.getValues();

  // Get the dimensions of the source range
  var numRows = sourceRange.getNumRows();
  var numCols = sourceRange.getNumColumns();

  // Paste data to target sheet
  targetSheet.getRange(2, 1, numRows, numCols).setValues(data);
}
