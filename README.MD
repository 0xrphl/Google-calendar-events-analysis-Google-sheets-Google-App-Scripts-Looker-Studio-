# Google Calendar to Google Sheets Looker Studio Integration

## Overview

This project aims to extract event data from Google Calendar and store it in a Google Sheets spreadsheet using Google Apps Script. The extracted data includes information such as Event Title, Start Date, End Date, Event Creator, Last Update, First Created, Activity Log, and Calendar ID. This allows for better accountability by tracking who created events and their last modifications.

Additionally, the extracted data is utilized in Looker Studio dashboards to gain insights into all the consultations that took place in the office.

## Features

- **Google Calendar Integration**: Automatically fetches event data from Google Calendar.
- **Google Sheets Integration**: Stores the extracted event data in a Google Sheets spreadsheet.
- **Accountability Tracking**: Tracks event creators and their last modifications.
- **Looker Studio Dashboard Integration**: Utilizes the extracted data in Looker Studio dashboards to gain insights.

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
-[Watch Video Tutorial how to select and create events](https://youtu.be/yZUaAZnF0N4)

-[Watch Video Tutorial how to clone crm tool](https://youtu.be/jfdt8V68hUs)

## Getting Started

### Setup Instructions
1. **Clone the Google Sheet**: Make a copy of the [Google Sheet template](https://docs.google.com/spreadsheets/d/1_gCUcBLbO2BsPPunIUigJN8JzGDVv4WtyhBUIf-ewWI/edit?usp=sharing) provided.
2. **Enable Google Apps Script**: In the Google Sheet, go to `Extensions` > `Apps Script` and enable scripts.
3. **Set Up Triggers**:
    - Go to `Edit` > `Current project's triggers` in Apps Script.
    - Click on `Add Trigger`.
    - Select the following options:
        - Choose which function to run: `checkCalendar`
        - Choose which deployment should run: `Head`
        - Select event source: `From spreadsheet`
        - Select event type: `On edit`
4. **Authorize Access**: Authorize necessary permissions for the script to function properly.
5. **Configure Looker Studio Integration**: [Instructions on integrating with Looker Studio].

## Support
For any questions or issues, please contact [Your Contact Information].

## Explanation of checkCalendar() Function
The `checkCalendar()` function is responsible for checking Google Calendar for events created based on the data entered into the Google Sheet. Here's how it works:

```javascript
function checkCalendar() {
  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Google Calendar and SMS tool");
  var range = sheet.getDataRange();
  var values = range.getValues();
  var long = (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Google Calendar and SMS tool").getRange(2,9).getValue()) + 1;

  // Loop through each row of the spreadsheet
  for (var i = 4; i < long + 3; i++) {
    var row = i + 1;
    var date = values[i][6]; // Adjusted column index for date
    var name = values[i][3]; // Adjusted column index for name
    var calendarId = "";  // Default calendar ID

    var calendarValue = values[i][8]; // Adjusted column index for calendar

    // Determine the calendar ID based on the value in the specified column
    if (calendarValue == "New York Office") {
      calendarId = "6a430b5ee8f9ea5a8237e8fe82533d7c8505e896edbaa5880c86ef117830b933@group.calendar.google.com"; 
    } else if (calendarValue == "Dallas Office") {
      calendarId = "d69f1c80512b7def2cf17d9edf8e5da3c83eb44b334c7a7f050cf38640f8f244@group.calendar.google.com";
    } else if (calendarValue == "Boston Office") {
      calendarId = "53d4f063c5cbd84330fd1a7e18074754ad149afce51c04bf1472aaa17102aefa@group.calendar.google.com";
    } else if (calendarValue == "Virtual Consultation") {
      calendarId = "9a9705841364b5ad8216b3ad6ef461ac7747f0e27841e26de539244dfe219863@group.calendar.google.com";
    } 

    // Get events from the specified calendar for the given date and name
    var events = CalendarApp.getCalendarById(calendarId).getEventsForDay(new Date(date), {search: name});
    Logger.log('Number of events: ' + events.length + ' - ' + name + ' - ' + date);

    // Update spreadsheet with event status based on the presence of events
    if (events.length != 0) {
      sheet.getRange(row, 22).setValue('Event Created'); // Adjusted column index for status
      sheet.getRange(row, 21).setValue(false); // Adjusted column index for checkbox
    } else {
      sheet.getRange(row, 22).setValue('Creation Pending'); // Adjusted column index for status
    }
  }

  // Clear status column for rows beyond the length

```

## Explanation of createCalendarEvents() Function
The `createCalendarEvents()` function creates events in Google Calendar based on the data entered into the Google Sheet. Here's how it works:

```javascript
function createCalendarEvents() {
  var sheet = SpreadsheetApp.getActiveSheet(); // Get active sheet
  var range = sheet.getDataRange(); // Get the range of data
  var values = range.getValues(); // Get values from the range
  
  for (var i = 4; i < values.length; i++) { // Loop through each row starting from 5th row (index 4)
    var row = i + 1; // Current row number
    var tickBox = sheet.getRange(row, 21).getValue(); // Get value of tickbox in the row
    var calendarId = "c_e84f80024a2ad3b161a1a115d578b2eb0c2a6c42be2fffdf2e0d41fbcd76242b@group.calendar.google.com"; // Default calendar ID

    if (tickBox) { // If tickbox is checked
      var eventDate = values[i][6]; // Get event date from the row
      var eventTime = values[i][11]; // Get event time from the row
      var eventTitle = "Initial Consultation - " + values[i][10] + " - " + values[i][7]; // Generate event title
      var eventDescription = "-<b> Name of the caller(lead): </b> " + values[i][3] +  "\n" + ...; // Generate event description
      var calendarValue = values[i][8]; // Get calendar value from the row
      var dateString = Utilities.formatDate(eventDate, "GMT", "yyyy-MM-dd"); // Format event date as string
      var timeString = eventTime.getHours() + ":" + eventTime.getMinutes() + ":" + eventTime.getSeconds(); // Format event time as string
      var dateTimeString = dateString + " " + timeString; // Concatenate date and time strings
      var today = new Date(); // Get current date
      var start = new Date(dateTimeString); // Create start date object
      var end = new Date(start.getTime() + (30 * 60000)); // Create end date object
      var oneDayAhead = new Date(today); // Get date one day ahead
      oneDayAhead.setDate(oneDayAhead.getDate() + 1); // Set date one day ahead
      
      // Construct SMS message
      var to = values[i][4]; // Recipient phone number
      var body = ""; // SMS body
      var scheduled_at = new Date(eventDate.getTime()); // Get event date as timestamp
      scheduled_at.setTime(scheduled_at.getTime() - (60 * 60 * 1000)); // Adjust timestamp
      
      // Set calendar ID based on location
      if (calendarValue == "New York Office") {
        calendarId = "6a430b5ee8f9ea5a8237e8fe82533d7c8505e896edbaa5880c86ef117830b933@group.calendar.google.com"; 
        eventTitle += " - New York Office In person"; // Append location to event title
      } 
      // Additional if conditions for other office locations
      
      // If virtual consultation, modify event title and description
      if (calendarValue == "Virtual Consultation") {
        calendarId = "9a9705841364b5ad8216b3ad6ef461ac7747f0e27841e26de539244dfe219863@group.calendar.google.com"; 
        eventTitle += " - Virtual ZOOM"; // Append "Virtual ZOOM" to event title
        eventTitle += " - " + values[i][0]; // Append event ID to event title
        eventDescription +=  "\n" +  "\n" + "Zoom link https://us02web.zoom.us/j/example?pwd=example21432sd" +  "\n" + "Meeting ID: xxxyyyxxx12" +  "\n" + "Passcode: 225478"; // Add Zoom link to event description   
      } 

      // Send SMS messages based on language
      if (values[i][2] === "SPA") {
        // Construct SMS message in Spanish
      } 
      if (values[i][2] === "ENG") {
        // Construct SMS message in English
      }  
      
      eventDescription +=  "\n" +  "\n" +  "*Event Created By Virtual Assistant. " + values[i][9]; // Add creator information to event description
      var calendar = CalendarApp.getCalendarById(calendarId); // Get calendar by ID
      var event = calendar.createEvent(eventTitle, start, end, {description: eventDescription}); // Create event in calendar
      Logger.log('Event ID: ' + event.getId()); // Log event ID
      sheet.getRange(row, 21).setValue(false); // Reset tickbox value
      sheet.getRange(row, 22).setValue('Event Created'); // Set status to "Event Created"
    }
  }
}
```

## Explanation of sendTextMessage() and scheduleTextUsMessage() Function
The sendTextMessage() and scheduleTextUsMessage() sends the template created to the SMS API to send consultation confirmation and schedule reminders for the client

```javascript
// Function to send a text message using the TextUS API
function sendTextMessage(to, body) {
  var apiKey = 'TextUS API Token'; // Your TextUS API token
  var url = 'https://next.textus.com//messages'; // API endpoint for sending messages

  // Construct message data
  var messageData = {
    to: to, // Recipient phone number
    body: body // Message body
  };

  // Set options for the HTTP request
  var options = {
    method: 'post', // HTTP method
    contentType: 'application/json', // Content type
    headers: {
      'Accept': 'application/vnd.textus+jsonld', // Accept header
      'Authorization': 'Bearer ' + apiKey // Authorization header with API token
    },
    payload: JSON.stringify(messageData) // Convert message data to JSON string
  };

  try {
    // Make the HTTP request to send the message
    var response = UrlFetchApp.fetch(url, options);

    // Get response status code
    var statusCode = response.getResponseCode();

    // Log success message if status code indicates success
    if (statusCode == 201 || statusCode == 202) {
      Logger.log('Message sent successfully.');
    } else {
      // Log error message if status code indicates failure
      var responseText = response.getContentText();
      Logger.log('Failed to send message. Status code: ' + statusCode + ', Response: ' + responseText);
    }
  } catch (error) {
    // Log error if an exception occurs during the request
    Logger.log('Error sending message:', error);
  }
}

// Function to schedule a text message using the TextUS API
function scheduleTextUsMessage(to, body, scheduled_at) {
  // Set API endpoint
  var apiUrl = 'https://next.textus.com//messages';

  // Set authorization token
  var token = 'TextUS API Token';

  // Construct message body
  var messageBody = {
    to: to, // Recipient phone number
    body: body, // Message body
    scheduled_at: scheduled_at // Scheduled time for sending the message
  };

  // Convert message body to JSON string
  var payload = JSON.stringify(messageBody);

  // Set options for the HTTP request
  var options = {
    "method": "POST", // HTTP method
    "contentType": "application/json", // Content type
    "headers": {
      "Accept": "application/vnd.textus+jsonld", // Accept header
      "Authorization": "Bearer " + token // Authorization header with API token
    },
    "payload": payload // Payload containing the message data
  };

  // Make the HTTP request to schedule the message
  var response = UrlFetchApp.fetch(apiUrl, options);

  // Log the response
  Logger.log(response.getContentText());
}
