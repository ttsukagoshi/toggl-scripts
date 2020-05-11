
/**
 * Google App Script to check the iCalID of a specific Google Calendar event.
 * You need to have access to the Google Calendar, as designated by the Google Calendar ID
 */
function checkICalId() {
  var calendarId = '37ctak8egrcaceen08dhi8cbjc@group.calendar.google.com';
  var startTime = new Date('2019-09-25T18:27+0900');
  var endTime = new Date('2019-09-25T18:30+0900');
  var events = CalendarApp.getCalendarById(calendarId).getEvents(startTime, endTime);
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    var eventTitle = event.getTitle();
    var eventICalId = event.getId();
    Logger.log('Title: ' + eventTitle + ', iCalId: ' + eventICalId);
  }
}

function checkTimeEntry() {
  var timeEntryIds = [1311945904];
  for (var i in timeEntryIds) {
    var timeEntryId = timeEntryIds[i];
    Logger.log(TogglScript.getTimeEntry(timeEntryId));
  }
}
