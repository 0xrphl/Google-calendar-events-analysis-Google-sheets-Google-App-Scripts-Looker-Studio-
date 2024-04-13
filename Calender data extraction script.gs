function CalenderDataExtractionScript() {
   var sheetName = "Calendars data extraction tool";
  var cellToUpdate = "E2";
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // Get today's date
  var today = new Date();
  
  // Format date as MM/dd/yyyy
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy");
  
  // Set the date in the cell
  sheet.getRange(cellToUpdate).setValue(formattedDate);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.getRange("A5:11000").clear(); // Adjusted range for the new "Calendar" column and the variable name column

  var startDateCell = sheet.getRange("C2").getValue();
  var endDateCell = sheet.getRange("E2").getValue();

  if (!startDateCell || !endDateCell || !(startDateCell instanceof Date) || !(endDateCell instanceof Date)) {
    Logger.log("Invalid date format in C1 or E1 cell. Please enter valid dates (YYYY-MM-DD).");
    return;
  }

  var startDate = new Date(startDateCell);
  var endDate = new Date(endDateCell);
  endDate.setDate(endDate.getDate() + 1);

  var lastRow = sheet.getLastRow();
  var startRow = lastRow + 1; // Start appending from the next row after the last row

  var calendarIds = {
    "9a9705841364b5ad8216b3ad6ef461ac7747f0e27841e26de539244dfe219863@group.calendar.google.com": "Virtual Consultations Calendar",
    "53d4f063c5cbd84330fd1a7e18074754ad149afce51c04bf1472aaa17102aefa@group.calendar.google.com": "Boston Consultations Calendar",
    "d69f1c80512b7def2cf17d9edf8e5da3c83eb44b334c7a7f050cf38640f8f244@group.calendar.google.com": "Dallas Consultations Calendar",
    "6a430b5ee8f9ea5a8237e8fe82533d7c8505e896edbaa5880c86ef117830b933@group.calendar.google.com": "New York Consultations Calendar"
  };

  var data = [];
  var calendars = [CalendarApp.getCalendarById(Object.keys(calendarIds)[0]), CalendarApp.getCalendarById(Object.keys(calendarIds)[1]), CalendarApp.getCalendarById(Object.keys(calendarIds)[2]), CalendarApp.getCalendarById(Object.keys(calendarIds)[3])];

  calendars.forEach(function(calendar, index) {
    var calendarId = Object.values(calendarIds)[index];
    var events = calendar.getEvents(startDate, endDate);

    events.forEach(function(event) {
      var title = event.getTitle();
      var start = event.getStartTime();
      var end = event.getEndTime();
      var creator = event.getCreators().join(", ");
      var lastUpdated = event.getLastUpdated();
      var firstCreated = event.getDateCreated();
      var description = event.getDescription();

      data.push([title, start, end, creator, lastUpdated, firstCreated, description, calendarId]);
    });
  });
  data.sort(function(a, b) {
    return a[1] - b[1];
  });

  // Add sorted data to the sheet, starting from the second row (index 1)
  sheet.getRange(5, 1, data.length, data[0].length).setValues(data);
}
