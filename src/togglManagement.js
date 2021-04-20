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

/* global TogglScript */
/* exported checkApiToken, deleteApiToken, onOpen, saveApiToken */

// Name of spreadsheet worksheets
const SHEET_NAME_CONFIG = '00_Config';
const SHEET_NAME_CALENDAR_IDS = '01_Calendar IDs';
const SHEET_NAME_AUTO_TAG = '02_AutoTag';
const SHEET_NAME_SPREADSHEET_LIST = '10_Record Spreadsheet List';
// User property keys
const UP_KEY_API_TOKEN = 'togglToken';
const UP_KEY_LAST_TIME_ENTRY_ID = 'lastTimeEntryId';

/**
 * Spreadsheet menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Toggl')
    .addItem('Get Toggl Time Entries', 'togglRecord')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Setup')
        .addItem('Check Saved Token', 'checkApiToken')
        .addItem('Save API Token', 'saveApiToken')
        .addSeparator()
        .addItem('Delete Token', 'deleteApiToken')
    )
    .addSeparator() //////////////
    .addItem('Test', 'test') //////////////
    .addToUi();
}

/**
 * Check the user property for existing Toggl API token and return the value as an alert message on the spreadsheet.
 */
function checkApiToken() {
  var existingToken = PropertiesService.getUserProperties().getProperty(UP_KEY_API_TOKEN);
  var message = '';
  if (!existingToken) {
    message = 'No Toggl API Token saved. Save a new one from the spreadsheet menu "Toggl" > "Setup" > "Save API Token"';
  } else {
    message = `Toggl API Token: ${existingToken}\nTake care in handling this value; this is basically a set of ID and password.`;
  }
  console.info('[checkApiToken] Completed.');
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Save Toggle API Token in the user property.
 */
function saveApiToken() {
  var ui = SpreadsheetApp.getUi();
  var up = PropertiesService.getUserProperties();
  var existingToken = up.getProperty(UP_KEY_API_TOKEN);
  var canceledMessage = '[saveApiToken] Canceled.';
  try {
    if (existingToken) {
      // If token already exists in the user property, ask the user whether or not to proceed.
      let alertMessage = 'You already have a Toggl API Token saved in the user property of this script. Do you want to overwrite the token? You can check the existing token from the spreadsheet menu "Toggl" > "Setup" > "Check Saved Token"';
      let alertResponse = ui.alert(alertMessage, ui.ButtonSet.YES_NO);
      if (alertResponse !== ui.Button.YES) {
        throw new Error(canceledMessage);
      }
    }
    // Enter & save token on user property
    let tokenResponse = ui.prompt('Enter Toggl API Token to save in the user property of this script. See https://github.com/toggl/toggl_api_docs#api-token for information on where to find your token.', ui.ButtonSet.OK_CANCEL);
    if (tokenResponse.getSelectedButton() !== ui.Button.OK) {
      throw new Error(canceledMessage);
    }
    up.setProperty(UP_KEY_API_TOKEN, tokenResponse.getResponseText());
    let completeMessage = `[saveApiToken] Process completed: Your Toggl API Token has been saved in the user property of this script. No other accounts, including other accounts sharing this spreadsheet with you, will have access to this token.`;
    // Log & notify user
    console.info(completeMessage);
    ui.alert(completeMessage);
  } catch (error) {
    if (error.message !== canceledMessage) {
      console.error(error.stack);
    }
    ui.alert(error.stack);
  }
}

/**
 * Delete Toggl API Token saved in the user property.
 */
function deleteApiToken() {
  var ui = SpreadsheetApp.getUi();
  var canceledMessage = '[deleteApiToken] Canceled.';
  try {
    // Confirmation before proceeding
    let confirmationResponse = ui.alert('Deleting Toggl API Token from user property. Are you sure you want to continue?', ui.ButtonSet.YES_NO);
    if (confirmationResponse !== ui.Button.YES) {
      throw new Error(canceledMessage);
    }
    // Delete user property
    PropertiesService.getUserProperties().deleteProperty(UP_KEY_API_TOKEN);
    let completeMessage = `[deleteApiToken] Process completed: User property "${UP_KEY_API_TOKEN}" has been deleted.`;
    // Log & notify user
    console.info(completeMessage);
    ui.alert(completeMessage);
  } catch (error) {
    if (error.message !== canceledMessage) {
      // Log as error only if the error message does not match the canceled message
      console.error(error.stack);
    }
    ui.alert(error.stack);
  }
}

/**
 * Retrieve Toggl time entries, record on Google Sheets, and transcribe the respective entries as Google Calendar events.
 * Note that there is a limitation to the number of time entries that can be retrieved at one TogglScript.getTimeEntries(startDateString, endDateString) call.
 * See the official document for more details:
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#get-time-entries-started-in-a-specific-time-range
 */
function togglRecord() {
  // Basic variables
  var myEmail = Session.getActiveUser().getEmail();
  var userName = myEmail.substring(0, myEmail.indexOf('@')); // the *** in ***@myDomain.com
  var userProperties = PropertiesService.getUserProperties();
  var togglScript = new TogglScript(userProperties.getProperty(UP_KEY_API_TOKEN));
  var upLastTimeEntryId = userProperties.getProperty(UP_KEY_LAST_TIME_ENTRY_ID); // the last retrieved time entry ID recorded on user property
  var now = new Date();
  // Get configuration values from the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var config = getSheetsInfo_(ss);
  // Get the target spreadsheet and relevant sheet objects
  var targetSpreadsheetUrl = config[SHEET_NAME_SPREADSHEET_LIST].reduce((latestRow, row) => {
    if (!Object.keys(latestRow).length || row.YEAR >= latestRow.YEAR) {
      latestRow = row;
    }
    return latestRow;
  }, {}).URL;
  var targetSpreadsheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl);
  var recordSheet = targetSpreadsheet.getSheetByName(config[SHEET_NAME_CONFIG].RECORD_SHEET_NAME);
  var lastTimeEntryId = (upLastTimeEntryId ? upLastTimeEntryId : getMax_(recordSheet, 1, 1));
  // Variables for logs
  var logTime = togglScript.togglFormatDate(now);
  var logSheet = targetSpreadsheet.getSheetByName(config[SHEET_NAME_CONFIG].LOG_SHEET_NAME);
  try {
    // Create objects for workspaces and projects with their IDs and names as keys and values, respectively.
    let workspaceObj = togglScript.getWorkspaces().reduce((obj, workspace) => {
      obj[workspace.id] = workspace.name;
      return obj;
    }, {});
    let projectObj = Object.keys(workspaceObj).reduce((obj, workspaceId) => {
      togglScript.getWorkspaceProjects(workspaceId)
        .forEach(project => obj[project.id] = project.name);
      return obj;
    }, {});
    // Get time entries
    let timeEntries = togglScript.getTimeEntries();
    // Array for new time entries
    let newTimeEntries = timeEntries.reduce((entries, timeEntry) => {
      // Skip time entries that
      // 1) have already been recorded on spreadsheet or 
      // 2) is currently running
      if (timeEntry.id > lastTimeEntryId && timeEntry.duration >= 0) {
        // Get names for items in IDs
        let workspaceName = workspaceObj[timeEntry.wid];
        let projectName = (timeEntry.pid ? projectObj[timeEntry.pid] : 'NA');
        // Convert array of tags into a comma-segmented string
        let tagStr = (timeEntry.tags ? timeEntry.tags.join() : '');
        // Convert date & times into time zone of the Toggl management spreadsheet
        let startLocal = togglScript.togglFormatDate(new Date(timeEntry.start), timeZone);
        let stopLocal = togglScript.togglFormatDate(new Date(timeEntry.stop), timeZone);
        // Record on Google Calendar
        let targetCalendarId = config[SHEET_NAME_CALENDAR_IDS][workspaceName].CALENDAR_ID; // Target Google Calendar ID
        let calendarTitle = gcalTitle(projectName, timeEntry.description); // Calendar event title; see below for gcalTitle()
        let calendarDesc = gcalDesc(timeEntry.id, workspaceName, tagStr); // Calendar description; see below for gcalDesc()
        // let event = CalendarApp.getCalendarById(targetCalendarId).createEvent(calendarTitle, new Date(startLocal), new Date(stopLocal), { description: calendarDesc });
        let iCalId = CalendarApp.getCalendarById(targetCalendarId)
          .createEvent(calendarTitle, new Date(startLocal), new Date(stopLocal), { description: calendarDesc })
          .getId();
        // Store the time entry properties in a formatted array to record as a row on spreadsheet
        entries.push([
          timeEntry.id, // Time Entry ID
          timeEntry.wid, // Workspace ID
          workspaceName,
          timeEntry.pid, // Project ID
          projectName,
          timeEntry.description,
          tagStr,
          startLocal,
          stopLocal,
          timeEntry.duration, // time entry duration in seconds. Contains a negative value if the time entry is currently running.
          timeEntry.uid, // User ID
          timeEntry.guid,
          timeEntry.billable,
          timeEntry.duronly,
          togglScript.togglFormatDate(new Date(timeEntry.at), timeZone), // Last modified in local time
          iCalId,
          logTime,
          targetCalendarId,
          '' // an empty field at the last for 'updateFlag'
        ]);
      }
      return entries;
    }, []);
    console.log(newTimeEntries);
    // Throw error if no new time entry is available
    if (!newTimeEntries.length) {
      throw new Error('No new time entry available');
    }
    // Record in current spreadsheet
    recordSheet.getRange(recordSheet.getLastRow() + 1, 1, newTimeEntries.length, newTimeEntries[0].length)
      .setValues(newTimeEntries);
    // Update UP_KEY_LAST_TIME_ENTRY_ID in user properties
    userProperties.setProperty(UP_KEY_LAST_TIME_ENTRY_ID, getMax_(recordSheet, 1, lastTimeEntryId));
    // Log
    logSheet.appendRow([logTime, userName, `Recorded: ${newTimeEntries.length} Toggl time entries`]);
  } catch (e) {
    // Log error
    logSheet.appendRow([logTime, userName, e.stack]);
    // Email notification
    let mailSub = `[Error] Recording Toggl Time Entry ${logTime}`;
    var mailBody = `${e.stack}\n\nToggl Management Spreadsheet: \n${ss.getUrl()}\n\nRecord Spreadsheet: \n${targetSpreadsheetUrl}`;
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
  var logSheet = targetSpreadsheet.getSheetByName('Log');
  var logText = '';
  var log = [];
  var logTimestamp = togglScript.togglFormatDate(now);

  try {
    // Throw exception if no tag is set.
    if (tag01 == null) {
      throw new Error('No tag set for autoTag().');
    }

    // Get latest time entries
    var timeEntries = togglScript.getTimeEntries();

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
    var updatedTimeEntries = togglScript.bulkUpdateTags(timeEntryIds, tags, 'add');

    // Log results
    logText = 'Updated: ' + timeEntryIds.length + ' time entry(ies) tagged by autoTag().\n' + JSON.stringify(updatedTimeEntries);
    log = [logTimestamp, userName, logText];
    logSheet.appendRow(log);

  } catch (e) {
    logText = 'Error: autoTag():\n' + togglScript.errorMessage(e);
    log = [logTimestamp, userName, logText];
    logSheet.appendRow(log);
  }
}

//***********************
// Background Function(s)
//***********************

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
 * @param {number} numCol Column number of target column starting from 1
 * @param {number} initialValue Initial integer to compare with; defaults to 0
 * @return {number} The largest number in target column
*/
function getMax_(sheet, numCol, initialValue = 0) {
  return sheet.getRange(2, numCol, sheet.getLastRow() - 1).getValues().reduce((max, cur) => Math.max(max, cur), initialValue);
}

/**
 * Retrieve contents of a designated spreadsheet in form of a JavaScript object.
 * Note that all sheets must have the first row as its header.
 * @param {Object} spreadsheet Spreadsheet object which can be retrieved by Google Apps Script methods like SpreadsheetApp.getActiveSpreadsheet()
 * @param {array} targetSheets [Optional] An array of sheet names from which to retrieve data and convert to object.
 * @returns {Object} Object in the following format: { sheetName: [{header1: value01, ..., headerN: value0N}, {header1: value11, ..., headerN: value1N}, ..., {header1: valueM1, ..., headerN: valueMN}] }
 */
function getSheetsInfo_(spreadsheet) {
  return spreadsheet.getSheets().reduce((obj, sheet) => {
    let sheetName = sheet.getName();
    let sheetValues = sheet.getDataRange().getValues();
    if (sheetName === SHEET_NAME_CONFIG) {
      sheetValues.shift(); // Remove header row
      // Create an object where row[0] and row[1] are key-value pairs
      obj[sheetName] = sheetValues.reduce((configs, row) => {
        configs[row[0]] = row[1];
        return configs;
      }, {});
    } else if (sheetName === SHEET_NAME_CALENDAR_IDS) {
      sheetValues.shift(); // Remove header row
      obj[sheetName] = sheetValues.reduce((calIdIndex, row) => {
        // Create an object with Toggl workspace names as its keys
        calIdIndex[row[0]] = { // row[0] is Toggl workspace name
          'WORKSPACE_ID': row[1],
          'CALENDAR_ID': row[2]
        };
        return calIdIndex;
      }, {});
    } else if (sheetName === SHEET_NAME_AUTO_TAG) {
      sheetValues.shift(); // Remove header row
      obj[sheetName] = sheetValues.reduce((autoTagIndex, row) => {
        // Create an object with Toggl workspace IDs as its keys,
        // and an array of Toggl tag names as their respective values.
        if (!autoTagIndex[row[0]]) { // row[0] is Toggl workspace ID
          autoTagIndex[row[0]] = [];
        }
        autoTagIndex[row[0]].push(row[1]); // row[1] is the name of the tag to add
        return autoTagIndex;
      }, {});
    } else if (sheetName === SHEET_NAME_SPREADSHEET_LIST) {
      let header = sheetValues.shift();
      obj[sheetName] = sheetValues.map(row => header.reduce((o, k, i) => {
        o[k] = row[i];
        return o;
      }, {}));
    }
    return obj;
  }, {});
}


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
 */

////////////////
// dailyCheck //
////////////////

/**
 * Check (daily) for the current date; 
 * If current year is not equal with the script property 'currentYear', a new spreadsheet will be created under a designated Google Drive folder.
 * The spreadsheet ID of this new spreadsheet will replace script property 'currentSpreadsheetId'
 */
function dailyCheck() {
  var now = new Date();
  var year = Utilities.formatDate(now, timeZone, 'yyyy');
  var log = [];
  var logTimestamp = TogglScript.togglFormatDate(now);
  var logText = '';
  try {
    if (year !== scriptProperties.currentYear || currentSpreadsheetId == null) {
      var togglFolder = DriveApp.getFolderById(togglFolderId);
      var spreadsheetName = 'Toggl' + year + ' (' + userName + ')';

      // Create spreadsheet in designated Google Drive folder; see below for details of createSpreadsheet() function.
      var createdSpreadsheetId = createSpreadsheet(togglFolder, spreadsheetName);
      var createdSpreadsheet = SpreadsheetApp.openById(createdSpreadsheetId);
      var createdSpreadsheetUrl = createdSpreadsheet.getUrl();

      // Format the created spreadsheet
      // Set sheet name & create a new sheet
      var recordSheet = createdSpreadsheet.getSheets()[0].setName('Toggl_Record');
      var logSheet = createdSpreadsheet.insertSheet(1).setName('Log');
      var sheets = [recordSheet, logSheet];
      // Create a header row in the sheet
      var header = [];
      var headerItems = [];
      // for recordSheet
      headerItems[0] = ['TIME_ENTRY_ID',
        'WORKSPACE_ID',
        'WORKSPACE',
        'PROJECT_ID',
        'PROJECT',
        'DESCRIPTION',
        'TAGS',
        'START',
        'STOP',
        'DURATION_SEC',
        'USER_ID',
        'GUID',
        'BILLABLE',
        'DURONLY',
        'LAST_MODIFIED',
        'iCalID',
        'TIMESTAMP',
        'CALENDAR_ID',
        'updateFlag'
      ];
      // for logSheet
      headerItems[1] = ['TIMESTAMP', 'USERNAME', 'LOG'];
      // Define header style
      var headerStyle = SpreadsheetApp.newTextStyle().setBold(true).build();

      for (var i = 0; i < sheets.length; i++) {
        var sheet = sheets[i];
        header[0] = headerItems[i]; // header must be two-dimensional array
        // Set header items and set text style
        var headerRange = sheet.getRange(1, 1, 1, headerItems[i].length)
          .setValues(header)
          .setHorizontalAlignment('center')
          .setTextStyle(headerStyle);
        // Freeze the first row
        sheet.setFrozenRows(1);
        // Delete empty columns
        sheet.deleteColumns(headerItems[i].length + 1, sheet.getMaxColumns() - headerItems[i].length);
        // Set vertical alignment to 'top'
        sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setVerticalAlignment('top');
      }

      // Update script properties
      var updatedProperties = {
        'currentYear': year,
        'currentSpreadsheetId': createdSpreadsheetId
      };
      if (currentSpreadsheetId !== null) {
        updatedProperties['prevSpreadsheetId'] = currentSpreadsheetId;
      }
      sp.setProperties(updatedProperties, false);

      // Log result
      logText = 'Created spreadsheet';
      log = [logTimestamp, userName, logText]; // one-dimensional array for appendRow()
      logSheet.appendRow(log);

      // Notification by email
      var stop = new Date();
      var executionTime = (stop - now) / 1000; // convert milliseconds to seconds
      var notification = 'New spreadsheet for Toggl record created at ' + createdSpreadsheetUrl + '\nScript execution time: ' + executionTime + ' sec';
      MailApp.sendEmail(myEmail, '[Toggl] New Spreadsheet for Toggl Record Created', notification);
    } else {
      // Log result
      logText = 'Checked: use current spreadsheet';
      log = [logTimestamp, userName, logText];
      currentSpreadsheet.getSheetByName('Log').appendRow(log);
    }
  } catch (e) {
    var thisScriptId = ScriptApp.getScriptId();
    var url = 'https://script.google.com/d/' + thisScriptId + '/edit';
    var body = TogglScript.errorMessage(e) + '\n\nCheck script at ' + url;
    MailApp.sendEmail(myEmail, '[Toggl] Error in Daily Date Check', body)
  }
}

//***********************
// Background Function(s)
//***********************

/**
 * Function to create a Google Spreadsheet in a particular Google Drive folder
 * @param {Object} targetFolder - Google Drive folder object in which you want to place the spreadsheet
 * @param {string} ssName - name of spreadsheet 
 * @return {string} ssId - spreadsheet ID of created spreadsheet
 */
function createSpreadsheet(targetFolder, ssName) {
  var ssId = SpreadsheetApp.create(ssName).getId();
  var temp = DriveApp.getFileById(ssId);
  targetFolder.addFile(temp);
  DriveApp.getRootFolder().removeFile(temp);
  return ssId;
}



////////////
// Update //
////////////

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
      var updateFlag = record[sheetLastCol - 1];
      // the ones digit of updateFlag
      var flagDet = parseInt(updateFlag) % 10;
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
      var duration = (stopM - startM) / 1000; // Convert milliseconds into seconds
      // For using PUT request on Toggl API
      var startToggl = TogglScript.togglFormatDateUpdate(startM);
      var stopToggl = TogglScript.togglFormatDateUpdate(stopM);;
      // Time entry contents to update/create
      var payload = {
        'time_entry': {
          'wid': workspaceId,
          'pid': projectId,
          'description': description,
          'tags': tagArray,
          'start': startToggl,
          'stop': stopToggl,
          'duration': duration,
          'created_with': 'GoogleAppScript'
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
      var newEvent = CalendarApp.getCalendarById(updatedTargetCalendar).createEvent(updatedCalendarTitle, startM, stopM, { description: updatedCalendarDesc });
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
  } catch (e) {
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
function getMaxTimeEntryId(targetSpreadsheet) {
  targetSpreadsheet = targetSpreadsheet || currentSpreadsheet;
  var recordSheet = targetSpreadsheet.getSheetByName('Toggl_Record');
  var max = getMax_(recordSheet, 1, 0);
  return max;
}
