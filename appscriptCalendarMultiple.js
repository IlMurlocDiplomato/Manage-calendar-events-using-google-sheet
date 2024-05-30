// Global variable to store calendar IDs and corresponding sheet names
var calendars = [
  { calendarId: "CALENDAR_ID", sheetName: "NAME_OF_SHEET" },
  { calendarId: "CALENDAR_ID", sheetName: "NAME_OF_SHEET" }
];

// Function to add the menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Manage Events')
    .addItem('Import Events from Calendars', 'importEventsFromCalendars')
    .addToUi();
}

// Function to cache events for a specific calendar
function cacheEvents(calendarId) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var today = new Date();
  var twoYearsFromNow = new Date(today.getFullYear() + 2, today.getMonth(), today.getDate());
  var events = calendar.getEvents(today, twoYearsFromNow);
  var cache = {};
  events.forEach(function(event) {
    cache[event.getId()] = event;
  });
  return cache;
}

// Function to import events from all calendars to their respective sheets
function importEventsFromCalendars() {
  calendars.forEach(function(calendar) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(calendar.sheetName);
    var cache = cacheEvents(calendar.calendarId);
    var eventsData = [];
    var data = sheet.getDataRange().getValues();
    var existingEventIds = data.slice(1).map(function(row) { return row[0]; }); // Assuming event ID is in the first column
    var today = new Date();
    var twoYearsFromNow = new Date(today.getFullYear() + 2, today.getMonth(), today.getDate());

    var events = CalendarApp.getCalendarById(calendar.calendarId).getEvents(today, twoYearsFromNow);
    events.forEach(function(event) {
      if (!existingEventIds.includes(event.getId())) {
        eventsData.push([event.getId(), event.getTitle(), event.getCreators().join(", "), Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd"), Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "HH:mm"), Utilities.formatDate(event.getEndTime(), Session.getScriptTimeZone(), "HH:mm"), event.getDescription().replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim()]);
      }
    });

    if (eventsData.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, eventsData.length, eventsData[0].length).setValues(eventsData);
    }
  });
}

// Function to handle cell edits
function onEdit(e) {
  var editedRange = e.range;
  var column = editedRange.getColumn();
  var sheetName = editedRange.getSheet().getName();

  calendars.forEach(function(calendar) {
    if (sheetName == calendar.sheetName) {
      var sheet = editedRange.getSheet();
      var row = editedRange.getRow();
      var eventId = sheet.getRange(row, 1).getValue();
      var endDate = sheet.getRange(row, 5).getValue();
      var startTimeString = sheet.getRange(row, 6).getValue();
      var endTimeString = sheet.getRange(row, 7).getValue();

      // Check if the edited column is E (end date)
      if (column == 5) {
        // Check if endDate is not null or empty
        if (endDate) {
          updateEventDate(calendar.calendarId, eventId, endDate);
        }
      }

      // Check if the edited column is F or G (start time or end time)
      if (column == 6 || column == 7) {
        // Check if startTimeString and endTimeString are not null or empty
        if (startTimeString && endTimeString) {
          var startTime = new Date('1970-01-01T' + startTimeString + ':00');
          var endTime = new Date('1970-01-01T' + endTimeString + ':00');
          // Check if startTime is before endTime
          if (startTime < endTime) {
            updateEventTime(calendar.calendarId, eventId, startTimeString, endTimeString);
          } else {
            Logger.log("Start time must be before end time.");
          }
        }
      }
    }
  });
}

// Function to update the event date
function updateEventDate(calendarId, eventId, endDate) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var event = calendar.getEventById(eventId);
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
    Logger.log("Event not found in the calendar.");
  }
}

// Function to update the event start and end times
function updateEventTime(calendarId, eventId, startTimeString, endTimeString) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var event = calendar.getEventById(eventId);
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
    Logger.log("Event not found in the calendar.");
  }
}
