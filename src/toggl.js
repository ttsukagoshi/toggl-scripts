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


/** List of script properties to set before executing script:
 * togglFolderId - ID of Google Drive folder to save Toggl spreadsheet in. 
 * currentSpreadsheetId - Current spreadsheet ID to record Toggl time entries in.
 * prevSpreadsheetId - [Optional] Spreadsheet ID of the spreadsheet used before currentSpreadsheetId.
 * togglToken - Toggl API Token. See https://github.com/toggl/toggl_api_docs#api-token for details.
 * calendarIdPrivate - Google Calendar ID for the calendar you want to save your time entries in workspace 'Private'
 * calendarIdWork - Google Calendar ID for the calendar you want to save your time entries in workspace 'Work'
 * lastTimeEntryId - the last retrieved time entry ID recorded on script property; togglRecord() will retrieve time entries that are larger than this ID.
 * currentYear - Current year.
 * autoTagWorkspaceId - [Optional] Target workspace ID of function autoTag()
 * autoTag01, autoTag02, ..., autoTag[n] - [Optional] Tags to use in function autoTag()
 */

/** 
 * Libary 'TogglScript' must be added to execute this script.
 * Script ID: 1gu8VZ-7Q1KdYpdIkUsijn_JiP6reGwUC2czBZtACzaJRyWxneM3MvXYY
 * See https://github.com/ttsukagoshi/Toggl_Scripts for details
 */

// Global variables
var sp = PropertiesService.getScriptProperties(); // File > Properties > Script Properties
var scriptProperties = sp.getProperties();
var togglFolderId = scriptProperties.togglFolderId;
var currentSpreadsheetId = scriptProperties.currentSpreadsheetId;
var currentSpreadsheet = SpreadsheetApp.openById(currentSpreadsheetId);
var prevSpreadsheetId = scriptProperties.prevSpreadsheetId;
var prevSpreadsheet = SpreadsheetApp.openById(prevSpreadsheetId);

var myEmail = Session.getActiveUser().getEmail();
var userName = myEmail.substring(0, myEmail.indexOf('@')); // the *** in ***@myDomain.com
var timeZone = currentSpreadsheet.getSpreadsheetTimeZone();

// Declare TogglScript Properties
// Toggl API Token. See https://github.com/toggl/toggl_api_docs#api-token for details.
TogglScript.togglToken = scriptProperties.togglToken; 
// Sync time zone of TogglScript Library to the zone of current spreadsheet
TogglScript.timeZone = timeZone;

// Create object for workspace name and its corresponding Google Calendar ID
var calendarIds = {};
// Change key name to fit your Toggl workspace name
// Google Calendar Id(s) must be set in the script properties beforehand.
// e.g., calendarIds['myWorkspaceName1'] = scriptProperties.calendarId****;
//       calendarIds['myWorkspaceName2'] = scriptProperties.calendarId****;
calendarIds['Private'] = scriptProperties.calendarIdPrivate;
calendarIds['Work'] = scriptProperties.calendarIdWork;

/**
 * Retrieve Toggl time entries, record on Google Spreadsheet, and transcribe to Google Calendar.
 * Set periodical triggers to execute this function. Note that there is a limitation
 * to the number of time entries that can be retrieved at one TogglScript.getTimeEntries(startDateString, endDateString) call.
 * See below documents for more details:
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#get-time-entries-started-in-a-specific-time-range
 * https://github.com/ttsukagoshi/Toggl_Scripts/blob/master/_TogglScriptLibrary.gs
 */
function togglRecord() {
  var now = new Date();
  var logTime = TogglScript.togglFormatDate(now);
  var currentSpreadsheetUrl = currentSpreadsheet.getUrl();
  var recordSheet = currentSpreadsheet.getSheetByName('Toggl_Record');
  var logSheet = currentSpreadsheet.getSheetByName('Log');
  var log = [];
  var logText = '';
  var lastTimeEntryId = scriptProperties.lastTimeEntryId; // the last retrieved time entry ID recorded on script property
  
  // Create an index object for workspace and project; see details of wpIndex() below
  var index = wpIndex();
  var workspaceIndex = index.workspaces;
  var projectIndex = index.projects;
  
  // Get time entries
  var timeEntries = TogglScript.getTimeEntries();

  // Array for new time entries
  var newEntries = [];
  var entryNum = 0; // resetting index for new time entries

  try {  
    for (var i = 0; i < timeEntries.length; i++) {
      var timeEntry = timeEntries[i];
      
      // Properties of a time entry; see https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md for details
      var timeEntryId = timeEntry.id;
      var workspaceId = timeEntry.wid;
      var projectId = timeEntry.pid;
      var desc = timeEntry.description;
      var tags = timeEntry.tags; // Array of tags in this time entry
      var start = timeEntry.start; // String. Time entry start time in ISO 8601 date and time.
      var stop = timeEntry.stop; // String . Time entry stop time in ISO 8601 date and time.
      var duration = timeEntry.duration; // time entry duration in seconds. Contains a negative value if the time entry is currently running.
      var userId = timeEntry.uid;
      var guId = timeEntry.guid;
      var billable = timeEntry.billable;
      var duronly = timeEntry.duronly;
      var lastModified = timeEntry.at;
      
      // Ignore time entries that 1) have already been recorded on spreadsheet or 2) is currently running
      if (timeEntryId <= lastTimeEntryId || duration < 0) {
        continue;
      }
      
      // Get names for items in IDs
      var workspaceName = workspaceIndex[workspaceId];
      var projectName = 'NA';
      if (projectId == null) {
        projectName = projectName;
      } else {
        projectName = projectIndex[projectId];
      }
      
      // Convert array of tags into a comma-segmented string
      var tag = '';
      if (tags == null){
        tag = tag;
      } else {
        tag = tags.join();
      }
      
      // Convert date & times into time zone of the current spreadsheet
      var startLocal = TogglScript.togglFormatDate(new Date(start));
      var stopLocal = TogglScript.togglFormatDate(new Date(stop));
      var lastModifiedLocal = TogglScript.togglFormatDate(new Date(lastModified));
      
      // Record on Google Calendar
      var targetCalendar = calendarIds[workspaceName]; // Target Google Calendar ID
      var calendarTitle = gcalTitle(projectName, desc); // Calendar event title; see below for gcalTitle()
      var calendarDesc = gcalDesc(timeEntryId, workspaceName, tags); // Calendar description; see below for gcalDesc()
      
      var event = CalendarApp.getCalendarById(targetCalendar).createEvent(calendarTitle, new Date(startLocal), new Date(stopLocal), {description: calendarDesc});
      var iCalId = event.getId();
      
      // Store into array newEntry to record into spreadsheet.
      var timestampLocal = logTime;
      var newEntry = [
        timeEntryId,
        workspaceId,
        workspaceName,
        projectId,
        projectName,
        desc,
        tag,
        startLocal,
        stopLocal,
        duration,
        userId,
        guId,
        billable,
        duronly,
        lastModifiedLocal,
        iCalId,
        timestampLocal,
        targetCalendar,
        '']; // an empty field at the last for 'updateFlag'
      newEntries[entryNum] = newEntry;
      entryNum += 1;  
    }
    
    // Throw error if no new time entry is available
    if (newEntries.length < 1) {
      throw new Error('No new time entry available');
    }
    
    // Record in current spreadsheet
    var recordRange = recordSheet.getRange(recordSheet.getLastRow()+1, 1, newEntries.length, recordSheet.getLastColumn()).setValues(newEntries);
    
    // Update lastTimeEntryId in script properties
    var uLastTimeEntryId = getMax(recordSheet, 1, lastTimeEntryId); // See below for details of getMax()
    sp.setProperty('lastTimeEntryId', uLastTimeEntryId);

    // Log
    logText = 'Recorded: ' + newEntries.length + ' Toggl time entries';
    log = [logTime, userName, logText];
    logSheet.appendRow(log);
  } catch (e) {
    logText = TogglScript.errorMessage(e);
    // Log error
    log = [logTime, userName, logText];
    logSheet.appendRow(log);
    // Email notification
    var mailSub = '[Error] Recording Toggl Time Entry ' + logTime;
    var thisScriptId = ScriptApp.getScriptId();
    var scriptUrl = 'https://script.google.com/d/' + thisScriptId + '/edit';
    var mailBody = logText + '\n\nScript: \n' + scriptUrl + '\n\nRecord Spreadsheet: \n' + currentSpreadsheetUrl;
    MailApp.sendEmail(myEmail, mailSub, mailBody);
  }
}

/**
 * Set particular tag(s) to all time entries in a workspace
 * e.g., set tag of the name of your office to all time entries in workspace 'Work'
 */
function autoTag() {
  var now = new Date();
  // Target workspace ID; if not specified in script property, all time entries will be subject to update.
  var targetWorkspaceId = scriptProperties.autoTagWorkspaceId || null;
  // Tags to add
  var tag01 = scriptProperties.autoTag01 || null;
  var tag02 = scriptProperties.autoTag02 || null;
  // var tag03 = scriptProperties.autoTag03, ..., tag** = scriptProperties.autoTag**;
  var tags = [tag01, tag02];

  var lastTimeEntryId = scriptProperties.lastTimeEntryId; // the last retrieved time entry ID recorded on script property
  var timeEntryIds = []; // Array of time entry IDs to update
  
  // Log
  var logSheet = currentSpreadsheet.getSheetByName('Log');
  var logText = '';
  var log = [];
  var logTimestamp = TogglScript.togglFormatDate(now);
  
  try {
    // Throw exception if no tag is set.
    if (tag01 == null) {
      throw new Error('No tag set for autoTag().');
    }
    
    // Get latest time entries
    var timeEntries = TogglScript.getTimeEntries();
    
    // Determine the time entries to update, i.e., time entries in designated workspace that are not recorded on the spreadsheet yet
    for (var i = 0; i < timeEntries.length; i++) {
      var timeEntry = timeEntries[i];
      var timeEntryId = timeEntry.id;
      var workspaceId = timeEntry.wid;
      var duration = timeEntry.duration; // time entry duration in seconds. Contains a negative value if the time entry is currently running.
      
      // Ignore time entries that 1) have already been recorded on spreadsheet, 2) is currently running, or 3) are not in the designated workspace
      if (timeEntryId <= lastTimeEntryId || duration < 0) {
        continue;
      } else if (targetWorkspaceId !== null && workspaceId !== targetWorkspaceId) {
        continue;
      } else {
        timeEntryIds.push(timeEntryId);
      }
    }
    
    // Throw exception if no time entry is subject to autoTag()
    if (timeEntryIds.length == 0) {
      throw new Error('No time entry subject to autoTag()');
    }
    
    // Bulk update time entries tags
    var updatedTimeEntries = TogglScript.bulkUpdateTags(timeEntryIds, tags, 'add'); 
  
    // Log results
    logText = 'Updated: ' + timeEntryIds.length + ' time entry(ies) tagged by autoTag().\n' + JSON.stringify(updatedTimeEntries);
    log = [logTimestamp, userName, logText];
    logSheet.appendRow(log);
    
  } catch(e) {
    logText = 'Error: autoTag():\n' + TogglScript.errorMessage(e);
    log = [logTimestamp, userName, logText];
    logSheet.appendRow(log);
  }
}

//***********************
// Background Function(s)
//***********************
/**
 * Create an index object for workspace and project, where
 * workspaceIndex = {workspaceId1=workspaceName1, workspaceId2=workspaceName2, ...} and
 * projectIndex = {projectId1=projectName1, projectId2=projectName2, ...}
 * @return {Object} {'workspaces'=workspaceIndex, 'projects'=projectIndex}
 */
function wpIndex() {
  var index = {};
  var workspaceIndex = {};
  var projectIndex = {};
  var workspaces = TogglScript.getWorkspaces();
  var wid = 0;
  var wname = '';
  var projects = [];
  var pid = 0;
  var pname = '';
  for (var i = 0; i < workspaces.length; i++) {
    wid = workspaces[i].id;
    wname = workspaces[i].name;
    workspaceIndex[wid] = wname;
    projects = TogglScript.getWorkspaceProjects(wid);
    for (var j = 0; j < projects.length; j++) {
      pid = projects[j].id;
      pname = projects[j].name;
      projectIndex[pid] = pname;
    }
  }
  index['workspaces'] = workspaceIndex;
  index['projects'] = projectIndex;
  return index;
}

/**
 * Standardized Google Calendar title format for this script
 *
 * @param {string} togglProjectName Toggl project name
 * @param {string} togglDesc Toggl description
 * @return {string} calendarTitle Google Calendar title 
 */
function gcalTitle(togglProjectName, togglDesc) {
  var calendarTitle = '[' + togglProjectName + '] ' + togglDesc;
  return calendarTitle;
}

/**
 * Standardized Google Calendar description format for this script
 * 
 * @param {number} timeEntryId Toggl time entry ID
 * @param {string} workspaceName Toggl workspace name
 * @param {string} tagsString tags in string for this Toggl time entry
 * @return {string} calendarDesc Google Calendar description
 */
function gcalDesc(timeEntryId, workspaceName, tagsString) {
  var calendarDesc = 'Time Entry ID: ' + timeEntryId + '\nWorkspace: ' + workspaceName + '\nTags: ' + tagsString;
  return calendarDesc;
}

/**
 * Returns the maximum value in a designated column
 * @param {sheet} sheet Target spreadsheet sheet object
 * @param {integer} numCol Column number of target column
 * @param {integer} initialValue Initial integer to compare with; defaults to zero
 * @return {integer} max The largest number in target column
*/
function getMax(sheet, numCol, initialValue) {
  initialValue = initialValue || 0;
  var data = sheet.getRange(2, numCol, sheet.getLastRow()-1).getValues();
  var max = data.reduce(function(accu, cur){return Math.max(accu, cur)}, initialValue);
  return max;
}

