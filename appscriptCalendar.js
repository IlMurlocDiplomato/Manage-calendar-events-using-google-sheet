// Global variables to store the calendar ID and cache
var calendarId = "CALENDAR_ID";
var eventCache = {}; // Cache to store events

// Function to add the menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Manage Events')
    .addItem('Import Events from Calendar', 'importEventsFromCalendar')
    .addToUi();
}

// Function to get all events from the calendar and store them in the cache
function cacheEvents() {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var today = new Date();
  var twoYearsFromNow = new Date(today.getFullYear() + 2, today.getMonth(), today.getDate());
  var events = calendar.getEvents(today, twoYearsFromNow);
  events.forEach(function(event) {
    eventCache[event.getId()] = event;
  });
}

// Function to update the event date
function updateEventDate(eventId, endDate) {
  var event = eventCache[eventId];
  if (event) {
    var startDate = new Date(endDate);
    var endDate = new Date(endDate);
    var endTime = event.getEndTime();
    var newEndDate = new Date(endDate);
    var newEndTime = new Date(endTime);
    newEndDate.setHours(newEndTime.getHours(), newEndTime.getMinutes(), newEndTime.getSeconds());

    event.setTime(startDate, newEndDate);
    Logger.log("Event date updated successfully.");
  } else {
    Logger.log("Event not found in the cache.");
  }
}

// Function to update the event start and end times
function updateEventTime(eventId, startTimeString, endTimeString) {
  var event = eventCache[eventId];
  if (event) {
    var eventStartDate = event.getStartTime();

    // Parse start time
    var startTimeParts = startTimeString.split(":");
    var newStartTime = new Date(eventStartDate);
    newStartTime.setHours(parseInt(startTimeParts[0], 10), parseInt(startTimeParts[1], 10));

    // Parse end time
    var endTimeParts = endTimeString.split(":");
    var newEndTime = new Date(eventStartDate);
    newEndTime.setHours(parseInt(endTimeParts[0], 10), parseInt(endTimeParts[1], 10));

    event.setTime(newStartTime, newEndTime);
    Logger.log("Event time updated successfully.");
  } else {
    Logger.log("Event not found in the cache.");
  }
}

// Function to import events from the calendar to the spreadsheet and update them
function importEventsFromCalendar() {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NAME_OF_SHEET");

  // Get the current date
  var today = new Date();

  // Calculate the date two years from now
  var twoYearsFromNow = new Date(today.getFullYear() + 2, today.getMonth(), today.getDate());

  // Get events from the calendar for the period between today and two years from now
  var events = calendar.getEvents(today, twoYearsFromNow);

  // Initialize an array to hold the data to be inserted into the spreadsheet
  var rowsToInsert = [];

  // Process the events in batch
  events.forEach(function(event) {
    var eventId = event.getId();
    var eventTitle = event.getTitle();
    var eventCreator = event.getCreators().join(", ");
    var eventDate = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var eventStartTime = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "HH:mm");
    var eventEndTime = Utilities.formatDate(event.getEndTime(), Session.getScriptTimeZone(), "HH:mm");

    // Clean the event description by removing HTML tags and ensuring it is on a single line
    var eventDescription = event.getDescription();
    var cleanedDescription = eventDescription ? eventDescription.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim() : "";

    // Check if the event already exists in the spreadsheet
    var eventExists = false;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) { // Assuming event ID is in the first column
        eventExists = true;
        break;
      }
    }

    // If the event does not already exist, add it to the array
    if (!eventExists) {
      rowsToInsert.push([eventId, eventTitle, eventCreator, eventDate, eventStartTime, eventEndTime, cleanedDescription]);
    }
  });

  // If there are events to insert, update or insert rows in the spreadsheet
  if (rowsToInsert.length > 0) {
    var startRow = sheet.getLastRow() + 1; // Determine the first empty row in the spreadsheet
    var endRow = startRow + rowsToInsert.length - 1; // Determine the last row to update

    // Update or insert rows in the spreadsheet
    sheet.getRange(startRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
  }
}

// Function to handle cell edits
function onEdit(e) {
  var editedRange = e.range;
  var column = editedRange.getColumn();
  var sheetName = editedRange.getSheet().getName();

  // Check if the edited sheet is "Calendar"
  if (sheetName == "Calendar") {
    var row = editedRange.getRow();
    var eventId = sheet.getRange(row, 1).getValue();
    var endDate = sheet.getRange(row, 8).getValue();
    var startTimeString = sheet.getRange(row, 9).getValue();
    var endTimeString = sheet.getRange(row, 10).getValue();

    // Check if the edited column is H (end date)
    if (column == 8) {
      // Check if endDate is not null or empty
      if (endDate) {
        updateEventDate(eventId, endDate);
      }
    }

    // Check if the edited column is I or J (start time or end time)
    if (column == 9 || column == 10) {
      // Check if startTimeString and endTimeString are not null or empty
      if (startTimeString && endTimeString) {
        var startTime = new Date('1970-01-01T' + startTimeString + ':00');
        var endTime = new Date('1970-01-01T' + endTimeString + ':00');
        // Check if startTime is before endTime
        if (startTime < endTime) {
          updateEventTime(eventId, startTimeString, endTimeString);
        } else {
          Logger.log("Start time must be before end time.");
        }
      }
    }
  }
}

// Cache events when the script starts
cacheEvents();
