# Couple of script for sync events of calendar, edit and confirm in case of collab
### They are just template to edit and use as you need. In this case I just show you how i use it so u can edit it in base at your need.
* appscriptCalendar.js it's made to manage one calendar in one page
* appscriptCalendarMultiple.js its made to manage more calendar in more page
### All calendar ar associeted with a single page of google sheets as we see

# Page preparation
I setted my documents with the follow columns
* ID (Is the evet id)
* Title (The title of the events)
* Creator (user who create the event)
* Proposed date (the date when the user who create the event propose)
* Proposed Start time (the start event time the user who create the event propose)
* Proposed Edn time (the end event time the user who create the event propose)
* Description (Decription of event)
* Confirmed date (the date that the editor "admin or pm" confirm or set for the event)
* Confirmed Start time (the start time that the editor "admin or pm" confirm or set for the event)
* Confirmed End time (the end time that the editor "admin or pm" confirm or set for the event)
## What this script do?
+ Import event from now on next 2 year to calendar when is open or when you click import event from calendar in the "Manage events" menu on the top
+ Change date, start time and end time of the event on the change of column Confirmed date,  Confirmed Start time,Confirmed End time
+ Clean the description, if the description contain some special tag or html element the script clean it and make it more readable
+ Check if modify have sense for save request
  + You cant set a end time that is before the start time
  + Dont change the time of event if you dont set Confirmed Start time AND Confirmed End time
+ Save time and api call
  + ALl the import request is made in batch
  + Use cache for save useless call
  + Evry "manual import" is refered only to the current Sheet
# Install
After you create your table you need to go in Extensione -> AppScript
Here you can paste the script. But first u need to make some change to make it compatible with your project.
## Change you need to do (this is valid for both file)
Replace ```CALENDAR_ID``` with your calendar id (you can find it in calendar settings)
Replace ```NAME_OF_SHEET``` whith the name you set to the sheet that manage the events
### On appscriptCalendar.js
On row 2:    ```var calendarId = "CALENDAR_ID";``` 

On row 68:   ```var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NAME_OF_SHEET");```
### On appscriptCalendar.js
At the begining you have a mapper
```
var calendars = [
  { calendarId: "CALENDAR_ID", sheetName: "NAME_OF_SHEET" },
  { calendarId: "CALENDAR_ID", sheetName: "NAME_OF_SHEET" }
];
```
Thats are just template and this is just an explanation for my personal use if you have different requirment you can still use this as template to work.


