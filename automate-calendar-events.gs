/**
 * The event handler triggered when opening the spreadsheet.
 * @param {Event} e The onOpen event.
 * @see https://developers.google.com/apps-script/guides/triggers#onopene
 */
var CALENDAR_NAME = "BBT LAB Events (automated)";

function triggerOpen(e) {
  // Set up Google Calendar
  var calendars = CalendarApp.getAllCalendars().map(x => x.getName());

  if ( ! calendars.includes(CALENDAR_NAME) ) {
    CalendarApp.createCalendar(CALENDAR_NAME);
  }
}

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */

var DATETIME_COLUMN = 1; // Column A
var LOCATION_COLUMN = 2; // Column B
var NAME_COLUMN     = 3; // Column C
var DESCRIPTION_COLUMN = 4; // Column D
var CALENDAR_COLUMN = 6; // Column F

function triggerEdit(e) {
  var row = e.range.getRow();
  var column = e.range.getColumn();
  var calendar_value = SpreadsheetApp.getActiveSheet().getRange(row, CALENDAR_COLUMN).getValue();
  
  // If row should be synced with calendar, create/update calendar event
  // Otherwise, delete calendar event if it doesn't exist
  if (shouldSyncCalendar(calendar_value)) {
    createOrUpdateCalendarEvent(row);
  } else {
    if (column == CALENDAR_COLUMN) {
      deleteCalendarEventIfExists(row);
    }
  }
}

function createOrUpdateCalendarEvent(row) {
  if (calendarEventExists(row)) {
    updateCalendarEvent(row);
  } else {
    createCalendarEvent(row);
  }
}

function deleteCalendarEventIfExists(row) {
  if (calendarEventExists(row)) {
    deleteCalendarEvent(row);
  }
}

// Calendar API calls
function createCalendarEvent(row) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var event_datetime = sheet.getRange(row, DATETIME_COLUMN).getValue();
  var event_location = sheet.getRange(row, LOCATION_COLUMN).getValue();
  var event_name = sheet.getRange(row, NAME_COLUMN).getValue();
  var event_description = sheet.getRange(row, DESCRIPTION_COLUMN).getValue();

  console.log("Creating event" + event_datetime, event_location, event_name, event_description);
  var calendar_event = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0].createEvent(event_name, new Date(event_datetime), new Date(event_datetime));

  // Save event id in note
  sheet.getRange(row, CALENDAR_COLUMN).setNote(calendar_event.getId());

  calendar_event.setTitle(event_name);
  calendar_event.setAllDayDate(new Date(event_datetime));
  calendar_event.setDescription(event_description);
  calendar_event.setLocation(event_location);
}

function updateCalendarEvent(row) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var event_datetime = sheet.getRange(row, DATETIME_COLUMN).getValue();
  var event_location = sheet.getRange(row, LOCATION_COLUMN).getValue();
  var event_name = sheet.getRange(row, NAME_COLUMN).getValue();
  var event_description = sheet.getRange(row, DESCRIPTION_COLUMN).getValue();

  console.log("Updating event" + event_datetime, event_location, event_name, event_description);
  var calendar_id = sheet.getRange(row, CALENDAR_COLUMN).getNote();
  var calendar_event = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0].getEventById(calendar_id);
  
  calendar_event.setTitle(event_name);
  calendar_event.setAllDayDate(new Date(event_datetime));
  calendar_event.setDescription(event_description);
  calendar_event.setLocation(event_location);
}

function deleteCalendarEvent(row) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var event_datetime = sheet.getRange(row, DATETIME_COLUMN).getValue();
  var event_location = sheet.getRange(row, LOCATION_COLUMN).getValue();
  var event_name = sheet.getRange(row, NAME_COLUMN).getValue();
  var event_description = sheet.getRange(row, DESCRIPTION_COLUMN).getValue();

  console.log("Deleting event" + event_datetime, event_location, event_name, event_description);
  var calendar_id = sheet.getRange(row, CALENDAR_COLUMN).getNote();
  var calendar_event = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0].getEventById(calendar_id);
  
  calendar_event.deleteEvent();
}

function calendarEventExists(row) {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Check if event was manually deleted
  var calendar = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0].getId();
  var calendar_id = sheet.getRange(row, CALENDAR_COLUMN).getNote();
  var eventId = calendar_id.split("@")[0];

  if(Calendar.Events.get(calendar, eventId).status === "cancelled") {
    console.log("Found cancelled event!");    
    return false; // Create a new one
  }
  
  return calendar_id != "";
}

// Helper functions
function shouldSyncCalendar(cellValue) {
  return (cellValue.toLowerCase() == "y" || cellValue.toLowerCase() == "yes");
}
