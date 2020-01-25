## Prepare
This is a stand-alone Google App Script.  
1. Create a script file on Google Drive
1. Register *Toggl Script* library (Script ID: `1gu8VZ-7Q1KdYpdIkUsijn_JiP6reGwUC2czBZtACzaJRyWxneM3MvXYY`)  
You can see the code for *Toggl Script* library at **_TogglScriptLibrary.gs**. The default identifier `TogglScript` is used in all scripts.
1. Copy & paste codes in **toggl.gs**`[required]`, **dailyCheck.gs**`[required]`, and **update.gs**`[optional]`. The purpose of each file is described below.
1. Set script properties as described in *List of script properties to set before executing script* of **toggl.gs**. Keep `currentSpreadsheetId` blank.
1. Execute function `dailyCheck()` in **dailyCheck.gs** once. Check to see that script property `currentSpreadsheetId` now has a value.

## Files
| File Name | Required/Optional | Purpose |
| --- | --- | --- |
| toggl.gs | Required | The main script for recording Toggl time entries on Google Spreadsheet and Calendar. Use triggers to periodically execute function `togglRecord()` |
| dailyCheck.gs | Required | Function `dailyCheck()` creates a new Google Spreadsheet at the change of the year for recording Toggle time entries for that year. Use to avoid hitting limitations in the maximum number of cells that can be created in a spreadsheet. |
| update.gs | Optional | Use function `updateToggl()` to update time entries. 1) Modify the time entry in spreadsheet, 2) add `1` to the integer in field `updateFlag`, and 3) execute function `updateToggl()` to change the original Toggl time entry(ies) and its corresponding Google Calendar event. Note that this function will always delete the original calendar event and create a new one; any changes made on the calendar event without updating the spreadsheet or the Toggl time entry itself will be lost.|
