// MIT License
// 
// Copyright (c) 2020 Taro TSUKAGOSHI
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

/**
 * Update Toggl and Google Calendar records for time entries where the value in spreadsheet for field 'updateFlag' is 1.
 */
function updateToggl() {
  var targetSpreadsheet = prevSpreadsheet; // Designate the spreadsheet object that you want to update

  // Stop currently running time entry
  var timestamp = TogglScript.togglFormatDate(new Date());
  var stoppedTimeEntry = TogglScript.stopRunningTimeEntry();
  var logText = 'Running Time Entry Stopped for Time Entry(ies) Update: \n' + JSON.stringify(stoppedTimeEntry);
  var log = currentSpreadsheet.getSheetByName('Log').appendRow([timestamp, userName, logText]);
  
  // Record all available time entries
  togglRecord();
  
  // Update the target time entries
  updateTimeEntries(targetSpreadsheet);
}

/**
 * Gets the appropriate value of lastTimeEntryId in designated spreadsheet and records it in the log.
 */
function getLastTimeEntryId() {
  var ssId = currentSpreadsheetId;
  var targetSpreadsheet = SpreadsheetApp.openById(ssId);
  var max = getMaxTimeEntryId(targetSpreadsheet);
  Logger.log('lastTimeEntryId should be ' + max);
}


//**************************
// Background Function(s)
//**************************
/**
 * Update Toggl and Google Calendar records for time entries
 * where the value in target spreadsheet for field 'updateFlag' is 1.
 * Numbers in updateFlag:
 * {'0': delete,
 *  ones digit = '1': update,
 *  hundreds digit : number of updates made
 * } 
 *
 * @param {Spreadsheet} targetSpreadsheet Spreadsheet object to update time entries. Defaults to the current spreadsheet.
*/
function updateTimeEntries(targetSpreadsheet) {
  targetSpreadsheet = targetSpreadsheet || currentSpreadsheet;
  var now = new Date();
  var timestampLocal = TogglScript.togglFormatDate(now);
  // Log
  var logSheet = targetSpreadsheet.getSheetByName('Log');
  var logText = '';
  var log = [];
  // Array to store the time entry before and after update
  var logOldNew = [];
  // Array to store updated time entry IDs
  var updatedIds = [];
  // Starting column number of columns to modify after time entry is updated 
  var updateRangeStartCol = 15;
  var updates = [];
  
  // Get the time entries in the spreadsheet
  var recordSheet = targetSpreadsheet.getSheetByName('Toggl_Record');
  var sheetLastCol = recordSheet.getLastColumn();
  var recordRange = recordSheet.getRange(2, 1, recordSheet.getLastRow() - 1, sheetLastCol);
  var records = recordRange.getValues();
  
  try {
    // Update Toggl and Google Calendar records for time entries where the (ones digit) value for field 'updateFlag' is 1
    for (var i = 0; i < records.length; i++) {
      var record = records[i];
      // Array to store the time entry before and after update
      var logOldNewObj = {};
      var logOldNewString = '';
      
      // the last field in record = 'updateFlag'
      var updateFlag = record[sheetLastCol-1];
      // the ones digit of updateFlag
      var flagDet = parseInt(updateFlag)%10;
      // Skip records whose updateFlag is not 1
      if (flagDet !== 1) {
        continue;
      }
      
      // Content of time entry to update
      var timeEntryId = record[0];
      var workspaceId = record[1];
      var workspaceName = record[2];
      var projectId = record[3];
      var projectName = record[4];
      var description = record[5];
      var tagString = record[6]; // comma-segmented string of tags
      var start = record[7];
      var stop = record[8];
      var oldICalId = record[15];
      var oldTargetCalendar = record[17];
      
      // Formatting components of PUT or POST request on Toggl API
      // Convert tagString to an array
      var tagArray = tagString.split(',');
      // Format start and stop
      // Date objects of start and stop
      var startM = new Date(start);
      var stopM = new Date(stop);
      var duration = (stopM - startM)/1000; // Convert milliseconds into seconds
      // For using PUT request on Toggl API
      var startToggl = TogglScript.togglFormatDateUpdate(startM);
      var stopToggl = TogglScript.togglFormatDateUpdate(stopM);;
      // Time entry contents to update/create
      var payload = {
        'time_entry' : {
          'wid' : workspaceId,
          'pid' : projectId,
          'description' : description,
          'tags' : tagArray,
          'start' : startToggl,
          'stop' : stopToggl,
          'duration' : duration,
          'created_with' : 'GoogleAppScript'
        }
      };
      var payloadString = JSON.stringify(payload);

      // Original time entry
      var oldTimeEntry = TogglScript.getTimeEntry(timeEntryId);
      
      var newTimeEntry = {};
      // Update or create new time entry depending on whether there is a change in workspace or not
      if (oldTimeEntry.data.wid == undefined || oldTimeEntry.data.wid == null) {
        throw new Error('Workspace ID in oldTimeEntry not available; could not complete update.');
      } else if (workspaceId == oldTimeEntry.data.wid) {
        // Update time entry
        newTimeEntry = TogglScript.updateTimeEntry(timeEntryId, payloadString); 
      } else {
        // Create new time entry on new workspace and log result.
        newTimeEntry = TogglScript.createTimeEntry(payloadString);
        logText = 'Created: 1 Toggl time entry.\n' + JSON.stringify(newTimeEntry);
        log = [timestampLocal, userName, logText];
        logSheet.appendRow(log);
        // Delete original time entry and log result.
        TogglScript.deleteTimeEntry(timeEntryId);
        logText = 'Deleted: 1 Toggl time entry.\n' + JSON.stringify(oldTimeEntry);
        log = [timestampLocal, userName, logText];
        logSheet.appendRow(log);
        // Update script and spreadsheet for the time entry ID and lastTimeEntryId
        timeEntryId = newTimeEntry.data.id;
        recordSheet.getRange(i + 2, 1).setValue(timeEntryId);
        sp.setProperty('lastTimeEntryId', getMaxTimeEntryId(targetSpreadsheet))
      }
      
      // Updated 'description' for the Toggl time entry
      var updatedDescription = newTimeEntry.data.description;
      
      // Updated 'lastModified' timestamp in Toggl
      var updatedLastModified = new Date(newTimeEntry.data.at);
      
      // Updated components of Google Calendar event
      var updatedTargetCalendar = calendarIds[workspaceName]; // Target Google Calendar ID
      var updatedCalendarTitle = gcalTitle(projectName, updatedDescription); // Calendar event title; see toggl.gs for gcalTitle()
      var updatedCalendarDesc = gcalDesc(timeEntryId, workspaceName, tagString); // Calendar description; see toggl.gs for gcalDesc()
      
      // Update calendar event (Delete old event and create a new one)
      CalendarApp.getCalendarById(oldTargetCalendar).getEventById(oldICalId).deleteEvent();
      var newEvent = CalendarApp.getCalendarById(updatedTargetCalendar).createEvent(updatedCalendarTitle, startM, stopM, {description: updatedCalendarDesc});
      var updatedICalId = newEvent.getId();
      
      // Update spreadsheet
      var lastModifiedLocal = TogglScript.togglFormatDate(updatedLastModified);
      updateFlag += 99; 
      updates[0] = [lastModifiedLocal, updatedICalId, timestampLocal, updatedTargetCalendar, updateFlag];
      var updateRange = recordSheet
      .getRange(i + 2, updateRangeStartCol, 1, sheetLastCol - updateRangeStartCol + 1)
      .setValues(updates);
      
      // Record old and new time entries for log
      logOldNewObj['old'] = oldTimeEntry;
      logOldNewObj['new'] = newTimeEntry;
      logOldNewString = JSON.stringify(logOldNewObj);
      logOldNew.push(logOldNewString);
      
      updatedIds.push(timeEntryId);
    }
    if (updatedIds.length < 1) {
      throw new Error('No updates');
    }
    
    // Log result
    logText = 'Updated: ' + updatedIds.length + ' Toggl time entry(ies).\n' + logOldNew.join('\n');
    log = [timestampLocal, userName, logText];
    logSheet.appendRow(log);
  } catch(e) {
    logText = TogglScript.errorMessage(e);
    log = [timestampLocal, userName, logText];
    logSheet.appendRow(log);
  }
}

/**
 * Gets the appropriate value of lastTimeEntryId in designated spreadsheet
 * 
 * @param {Spreadsheet} targetSpreadsheet Spreadsheet object to update time entries. Defaults to the current spreadsheet.
*/
function getMaxTimeEntryId(targetSpreadsheet){
  targetSpreadsheet = targetSpreadsheet || currentSpreadsheet;
  var recordSheet = targetSpreadsheet.getSheetByName('Toggl_Record');
  var max = getMax(recordSheet, 1, 0);
  return max;
}
