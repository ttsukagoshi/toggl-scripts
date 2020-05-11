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
        'currentYear' : year,
        'currentSpreadsheetId' : createdSpreadsheetId
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
      var executionTime = (stop - now)/1000; // convert milliseconds to seconds
      var notification = 'New spreadsheet for Toggl record created at ' + createdSpreadsheetUrl + '\nScript execution time: ' + executionTime + ' sec';
      MailApp.sendEmail(myEmail, '[Toggl] New Spreadsheet for Toggl Record Created', notification);
    } else {
      // Log result
      logText = 'Checked: use current spreadsheet';
      log = [logTimestamp, userName, logText];
      currentSpreadsheet.getSheetByName('Log').appendRow(log);
    }
  } catch(e) {
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
