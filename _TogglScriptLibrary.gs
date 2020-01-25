/**
 * This is a Google App Script library for using Toggl API on app script
 * See official document at https://github.com/toggl/toggl_api_docs for details
 */

// Properties
/**
 * @properties {string} togglToken Toggl API token; to be declared at individual scripts
 * @properties {string} timeZone Time zone of script using this library; to be declared at individual scripts. Defaults to the time zone of this script.
 * @properties {string} apiVersion The basic URL with API version. https://github.com/toggl/toggl_api_docs/blob/master/toggl_api.md
 * @properties {string} _reportsApiVersion [Experimental] The basic URL of Toggl Reports API. https://github.com/toggl/toggl_api_docs/blob/master/reports.md
 * @properties {string} _myEmail [Experimental] Required when making GET requests on Reports API; to be declared at individual scripts. https://github.com/toggl/toggl_api_docs/blob/master/reports.md#request-parameters
 */
var togglToken = 'myToken';
var timeZone = Session.getScriptTimeZone();
var apiVersion = 'https://www.toggl.com/api/v8';
var _reportsApiVersion = 'https://toggl.com/reports/api/v2'; 
var _myEmail = '';

// Methods
/**
 * Generate a string of API token for authentication using HTTP basic auth and API Token.
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/authentication.md#http-basic-auth-with-api-token
 *
 * @return {string} auth String of '[myApiToken]:api_token' used for basic auth
 */
function basicAuth() {
  var auth = togglToken + ':api_token';
  return auth;
}

/**
 * GET request on Toggl API
 * https://github.com/toggl/toggl_api_docs/blob/master/toggl_api.md
 * 
 * @param {string} path
 * @return {String} response JSON-encoded object; needs to be JSON.parse()ed before using as object
 */
function get(path) {
  var url = apiVersion + path;
  var options = {
    'method' : 'GET',
    'contentType': 'application/json',
    'headers': {"Authorization" : "Basic " + Utilities.base64Encode(basicAuth())},
    'muteHttpExceptions':false
  };
  var response = UrlFetchApp.fetch(url, options);
  return response;
}

/**
 * PUT request on Toggl API
 *
 * @param {string} path
 * @param {String} payloadString [Optional] JSON-encoded details of PUT request
 * @return {String} response JSON-encoded object; needs to be JSON.parse()ed before using as object
 */
function put(path, payloadString) {
  payloadString = payloadString || null;
  var url = apiVersion + path;
  var options = {
    'method' : 'PUT',
    'contentType' : 'application/json',
    'headers' : {"Authorization" : "Basic " + Utilities.base64Encode(basicAuth())},
    'muteHttpExceptions' : false
  };
  if (payloadString !== null) {
    options['payload'] = payloadString;
  }
  var response = UrlFetchApp.fetch(url, options);
  return response;
}

/**
 * POST request on Toggl API
 *
 * @param {string} path
 * @param {String} payload JSON-encoded details of PUT request
 * @return {String} response JSON-encoded object; needs to be JSON.parse()ed before using as object
 */
function post(path, payload) {
  var url = apiVersion + path;
  var options = {
    'method' : 'POST',
    'contentType' : 'application/json',
    'headers' : {"Authorization" : "Basic " + Utilities.base64Encode(basicAuth())},
    'muteHttpExceptions' : false,
    'payload' : payload,
  };
  var response = UrlFetchApp.fetch(url, options);
  return response;
}

/**
 * DELETE request on Toggl API
 *
 * @param {string} path
 * @return {string} response HTTP response; '200 OK' when successful
 */
function tDelete(path) {
  var url = apiVersion + path;
  var options = {
    'method' : 'DELETE',
    'contentType' : 'application/json',
    'headers' : {"Authorization" : "Basic " + Utilities.base64Encode(basicAuth())},
    'muteHttpExceptions' : false
  };
  var response = UrlFetchApp.fetch(url, options);
  return response;
}

/**
 * GET request for time entries.
 * When startDateString and endDateString is specified, gets the time entries for the period between start and end dates.
 * If not, returns the time entries in the last 9 days. The limit of returned time entries is 1000.
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#get-time-entries-started-in-a-specific-time-range
 *
 * @param {string} startDateString - ISO 8601 date and time strings; e.g., "2020-01-09T19:23:30+09:00"
 * @param {string} endDateString - ISO 8601 date and time strings
 * @return {Object} timeEntries
 */
function getTimeEntries(startDateString, endDateString) {
  startDateString = startDateString || undefined;
  endDateString = endDateString || undefined;
  var extUrl = '';
  if (startDateString == null && endDateString == null) {
    extUrl = extUrl;
  } else if (startDateString == null || endDateString == null) {
    var paramMissing = 'parameter endDateString missing';
    return paramMissing;
  } else {
    extUrl = '?start_date=' + encodeURIComponent(startDateString) + '&end_date=' + encodeURIComponent(endDateString);
  }
  var path = '/time_entries' + extUrl;
  var timeEntries = JSON.parse(get(path));
  return timeEntries;
}

/**
 * Get details of a specific time entry
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#get-time-entry-details
 *
 * @param {number} timeEntryId
 * @return {Object} timeEntry Details of the time entry
 */
function getTimeEntry(timeEntryId) {
  var path = '/time_entries/' + timeEntryId;
  var timeEntry = JSON.parse(get(path));
  return timeEntry;
}

/**
 * Get details of a running time entry
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#get-running-time-entry
 *
 * @return {Object} runningTimeEntry Details of the running time entry. Note that runningTimeEntry.data.duration is a negative value.
 */
function getRunningTimeEntry() {
  var path = '/time_entries/current';
  var runningTimeEntry = JSON.parse(get(path));
  return runningTimeEntry;
}

/**
 * Get all workspaces which belongs to the token owner.
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/workspaces.md#get-workspaces
 *
 * @return {Object} workspaces
 */
function getWorkspaces() {
  var path = '/workspaces';
  var workspaces = JSON.parse(get(path));
  return workspaces;
}

/**
 * Get details of the specified workspace.
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/workspaces.md#get-single-workspace
 *
 * @param {number} workspaceId
 * @return {Object} workspaces
 */
function getWorkspace(workspaceId) {
  var path = '/workspaces/' + workspaceId;
  var workspace = JSON.parse(get(path));
  return workspace;
}

/**
 * Get all project(s) in a specified workspace. To get a successful response, the token owner must be workspace admin.
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/workspaces.md#get-workspace-projects
 *
 * @param {number} workspaceId
 * @return {Object} workspaceProjects
 */ 
function getWorkspaceProjects(workspaceId) {
  var path = '/workspaces/' + workspaceId + '/projects';
  var workspaceProjects = JSON.parse(get(path));
  return workspaceProjects;
}

/**
 * Get details of a specified project.
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/projects.md#get-project-data
 *
 * @param {number} projectId
 * @return {Object} project
 */ 
function getProject(projectId) {
  var path = '/projects/' + projectId;
  var project = JSON.parse(get(path));
  return project;  
}

/**
 * POST request to create a time entry
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#create-a-time-entry
 * 
 * @param {String} payloadString JSON-encoded update details which should be in form of {'time_entry' : {'wid' : workspaceId, 'pid' : projectId, 'description' : description, 'tags' : tagArray, 'start' : startToggl, 'stop' : stopToggl, 'duration' : duration}
 * @return {Object} createdEntry The created time entry.
 */
function createTimeEntry(payloadString) {
  var path = '/time_entries';
  var createdEntry = JSON.parse(post(path, payloadString));
  return createdEntry;
}

/**
 * PUT request to stop a time entry
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#stop-a-time-entry
 *
 * @param {number} timeEntryId ID of the time entry to stop
 * @return {Object} stoppedTimeEntry The stopped time entry
 */
function stopTimeEntry(timeEntryId) {
  var path = '/time_entries/' + timeEntryId + '/stop';
  var stoppedTimeEntry = JSON.parse(put(path));
  return stoppedTimeEntry;
}

/**
 * PUT request to bulk update time entries tags
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#bulk-update-time-entries-tags
 *
 * @param {Array} timeEntryIds Array of time entry IDs to update tags
 * @param {Array} tags Array of tags in string
 * @param {string} tagAction [Optional] string of either 'add' or 'remove'. When this is left blank, update will override all existing tags on the time entries.
 * @return {Object} updatedTimeEntries The updated time entries
 */
function bulkUpdateTags(timeEntryIds, tags, tagAction) {
  var path = '/time_entries/' + timeEntryIds.join();
  var payload = {
    'time_entry' : {
      'tags' : tags,
      'tag_action' : tagAction
    }
  }
  var payloadString = JSON.stringify(payload);
  var updatedTimeEntries = JSON.parse(put(path, payloadString));
  return updatedTimeEntries;
}

/**
 * Stop the running time entry
 *
 * @return {Object} stoppedTimeEntry The stopped time entry
 */
function stopRunningTimeEntry() {
  var runningTimeEntryId = getRunningTimeEntry().data.id;
  var stoppedTimeEntry = stopTimeEntry(runningTimeEntryId);
  return stoppedTimeEntry;
}

/**
 * PUT request to update a time entry
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#update-a-time-entry
 *
 * @param {number} timeEntryId ID of the time entry to update
 * @param {String} payloadString JSON-encoded update details which should be in form of {'time_entry' : {'wid' : workspaceId, 'pid' : projectId, 'description' : description, 'tags' : tagArray, 'start' : startToggl, 'stop' : stopToggl, 'duration' : duration}
 * @return {Object} updatedEntry The updated time entry.
 */
function updateTimeEntry(timeEntryId, payloadString) {
  var path = '/time_entries/' + timeEntryId;
  var updatedEntry = JSON.parse(put(path, payloadString));
  return updatedEntry;
}

/**
 * DELETE request to delete a time entry
 * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#delete-a-time-entry
 *
 * @param {number} timeEntryId ID of the time entry to delete
 * @return {string} response HTTP response; '200 OK' if successful
 */
function deleteTimeEntry(timeEntryId) {
  var path = '/time_entries/' + timeEntryId;
  var response = tDelete(path);
  return response;
}
  
/**
 * Converts a string representing date & time into designated ISO 8601 format.
 * Used when retrieving time entries of a specific period and recording date on spreadsheet.
 * 
 * @param {Date} date - date object
 * @return {string} dateIso
 */
function togglFormatDate(date) {
  var dateIso = Utilities.formatDate(date, timeZone, "yyyy-MM-dd'T'HH:mm:ssXXX");
  return dateIso;
}

/**
 * Converts a string representing date & time into designated ISO 8601 format.
 * Used for updating time entry in Toggl
 * @param {Date} date - date object
 * @return {string} dateIso
 */
function togglFormatDateUpdate(date) {
  var dateIso = Utilities.formatDate(date, 'GMT', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  return dateIso;
}

/**
 * Standarized error message for this script
 * @param {Object} e - error object returned by try-catch
 * @return {string} message - standarized error message
 */
function errorMessage(e) {
  var message = 'Error : line - ' + e.lineNumber + '\n[' + e.name + '] ' + e.message + '\n' + e.stack;
  return message;
}

//********************************************
// Experimental Methods (Toggl Reports API)
//********************************************

/**
 * GET request on Toggl Reports API
 * https://github.com/toggl/toggl_api_docs/blob/master/reports.md
 */
function _getReports(path) {
  var url = _reportsApiVersion + path;
  var options = {
    'method' : 'GET',
    'contentType': 'application/json',
    'headers': {"Authorization" : "Basic " + Utilities.base64Encode(basicAuth())},
    'muteHttpExceptions':false
  };
  var response = UrlFetchApp.fetch(url, options);
  return response;
}

/**
 * Get detailed time entries via Reports API
 * https://github.com/toggl/toggl_api_docs/blob/master/reports/detailed.md
 *
 * @param {number} wid Workspace ID of the workspace whose data you want to access.
 * @param {string} since ISO 8601 date string in yyyy-MM-dd, e.g., "2020-01-09"
 * @param {number} page Optional; page number (integer) of the paged response. Defaults to 1.
 * @return {Object} Time Entries
 */
function _getReportsDetails(wid, since, page) {
  page = page || 1;
  var path = '/details/?user_agent=' + encodeURIComponent(myEmail) + '&workspace_id=' + wid + '&since=' + since + '&page=' + parseInt(page);
  var response = _getReports(path);
  var object = JSON.parse(response);
  return object;
}
