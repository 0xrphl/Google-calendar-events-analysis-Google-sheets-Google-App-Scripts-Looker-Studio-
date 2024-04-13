# Google Calendar to Google Sheets Looker Studio Integration

## Table of Contents
1. [Overview](#overview)
2. [Features](#features)
3. [Flow Chart](#flow-chart)
   - [Demo](#demo)
4. [Video Tutorials](#video-tutorials)
5. [Getting Started](#getting-started)
   - [Setup Instructions](#setup-instructions)
6. [Support](#support)
7. [Calendar Data Extraction Script](#calendar-data-extraction-script)
   - [Script Explanation](#script-explanation)
   - [Explanation of copyDataToSheet2() Function](#explanation-of-copydatatosheet2-function)

## Overview

This project aims to extract event data from Google Calendar and store it in a Google Sheets spreadsheet using Google Apps Script. The extracted data includes information such as Event Title, Start Date, End Date, Event Creator, Last Update, First Created, Activity Log, and Calendar ID. This allows for better accountability by tracking who created events and their last modifications.

Additionally, the extracted data is utilized in Looker Studio dashboards to gain insights into all the consultations that took place in the office.

## Features

- **Google Calendar Integration**: Automatically fetches event data from Google Calendar.
- **Google Sheets Integration**: Stores the extracted event data in a Google Sheets spreadsheet.
- **Accountability Tracking**: Tracks event creators and their last modifications.
- **Looker Studio Dashboard Integration**: Utilizes the extracted data in Looker Studio dashboards to gain insights.
- **Daily Script Trigger**: Automatically runs a Google Apps Script on a daily basis to extract data and feed the Looker Studio dashboard.

## Flow Chart
![Flow Chart](https://i.ibb.co/wrKRm63/diagrama-drawio-2.png)

### Demo
- **Google Sheet Demo**: [Google Sheet Demo](https://docs.google.com/spreadsheets/d/1liepeB7uZDuwKyq3wZjjSy8qKixTkjxBZEPDTHuq7YM/edit?usp=sharing)
- **Looker Studio Demo**: [Looker Studio Demo](https://lookerstudio.google.com/reporting/2a8bfdcc-5fe8-403f-bf28-92565e05c518)
- **Calendars**:
  - Calendar 1: [Link](https://calendar.google.com/calendar/embed?src=53d4f063c5cbd84330fd1a7e18074754ad149afce51c04bf1472aaa17102aefa%40group.calendar.google.com&ctz=America%2FNew_York)
  - Calendar 2: [Link](https://calendar.google.com/calendar/embed?src=d69f1c80512b7def2cf17d9edf8e5da3c83eb44b334c7a7f050cf38640f8f244%40group.calendar.google.com&ctz=America%2FNew_York)
  - Calendar 3: [Link](https://calendar.google.com/calendar/embed?src=6a430b5ee8f9ea5a8237e8fe82533d7c8505e896edbaa5880c86ef117830b933%40group.calendar.google.com&ctz=America%2FNew_York)
  - Calendar 4: [Link](https://calendar.google.com/calendar/embed?src=9a9705841364b5ad8216b3ad6ef461ac7747f0e27841e26de539244dfe219863%40group.calendar.google.com&ctz=America%2FNew_York)

## Video Tutorials
-[Watch Video Tutorial how to extract events data](https://youtu.be/1cKc1VzYlpI)

-[Watch Video Tutorial how to clone google calendar tool](https://youtu.be/eW7maJzxtHI)

## Getting Started

### Setup Instructions
1. **Clone the Google Sheet**: Make a copy of the [Google Sheet template](https://docs.google.com/spreadsheets/d/1liepeB7uZDuwKyq3wZjjSy8qKixTkjxBZEPDTHuq7YM/edit?usp=sharing) provided.
2. **Enable Google Apps Script**: In the Google Sheet, go to `Extensions` > `Apps Script` and enable scripts.
3. **Set Up Triggers**:
    - Go to `Edit` > `Current project's triggers` in Apps Script.
    - Click on `Add Trigger`.
    - Select the following options:
        - Choose which function to run: `checkCalendar`
        - Choose which deployment should run: `Time-driven`
        - Select event source: `Day timer`
        - Select event type: `select timeframe of your preference`
        - Repeat the step number 3 with the function `copyDataToSheet2`
4. **Authorize Access**: Authorize necessary permissions for the script to function properly.
5. **Configure Looker Studio Integration**: [Instructions on integrating with Looker Studio].

## Support
For any questions or issues, please contact [here](https://github.com/0xrphl/Google-calendar-events-analysis-Google-sheets-Google-App-Scripts-Looker-Studio-/issues/new).

# Calendar Data Extraction Script

This script is designed to extract calendar data from multiple Google calendars and populate a Google Spreadsheet with the extracted information.

## Script Explanation

```javascript
function CalendarDataExtractionScript() {
  // Set configuration variables
  var sheetName = "Calendars data extraction tool"; // Name of the target sheet
  var cellToUpdate = "E2"; // Cell to update with today's date
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); // Get the target sheet
  
  // Get today's date and format it
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy");
  
  // Set the formatted date in the specified cell
  sheet.getRange(cellToUpdate).setValue(formattedDate);
  
  // Clear existing data in the specified range
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.getRange("A5:11000").clear(); // Adjusted range for the new "Calendar" column and the variable name column

  // Fetch start and end dates from the sheet
  var startDateCell = sheet.getRange("C2").getValue();
  var endDateCell = sheet.getRange("E2").getValue();

  // Validate date format
  if (!startDateCell || !endDateCell || !(startDateCell instanceof Date) || !(endDateCell instanceof Date)) {
    Logger.log("Invalid date format in C1 or E1 cell. Please enter valid dates (YYYY-MM-DD).");
    return;
  }

  // Convert start and end dates to Date objects
  var startDate = new Date(startDateCell);
  var endDate = new Date(endDateCell);
  endDate.setDate(endDate.getDate() + 1);

  // Get the last row index and calculate the starting row for appending new data
  var lastRow = sheet.getLastRow();
  var startRow = lastRow + 1;

  // Define calendar IDs and names
  var calendarIds = {
    "9a9705841364b5ad8216b3ad6ef461ac7747f0e27841e26de539244dfe219863@group.calendar.google.com": "Virtual Consultations Calendar",
    "53d4f063c5cbd84330fd1a7e18074754ad149afce51c04bf1472aaa17102aefa@group.calendar.google.com": "Boston Consultations Calendar",
    "d69f1c80512b7def2cf17d9edf8e5da3c83eb44b334c7a7f050cf38640f8f244@group.calendar.google.com": "Dallas Consultations Calendar",
    "6a430b5ee8f9ea5a8237e8fe82533d7c8505e896edbaa5880c86ef117830b933@group.calendar.google.com": "New York Consultations Calendar"
  };

  // Array to hold extracted data
  var data = [];
  
  // Fetch events from each calendar and populate the data array
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

      // Push event data to the data array
      data.push([title, start, end, creator, lastUpdated, firstCreated, description, calendarId]);
    });
  });

  // Sort data array by start time
  data.sort(function(a, b) {
    return a[1] - b[1];
  });

  // Populate the spreadsheet with the sorted data
  sheet.getRange(5, 1, data.length, data[0].length).setValues(data);
}
```

## Explanation of copyDataToSheet2() Function
The `copyDataToSheet2()` function export every day all the events from the office calendars to the database that feeds the looker studio dashboard

```javascript
function copyDataToSheet2() {
  // Get active spreadsheet and source & target sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Calendars data extraction tool"); // Source sheet
  var targetSheet = ss.getSheetByName("Looker data"); // Target sheet

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
```