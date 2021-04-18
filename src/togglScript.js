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

/* exported TogglScript */

/**
 * This file defines the class for using Toggl API on Google Apps Script.
 * See official documents at https://github.com/toggl/toggl_api_docs for details.
 */

const TOGGL_API_VERSION = 'https://api.track.toggl.com/api/v8';
const _REPORTS_API_VERSION = 'https://api.track.toggl.com/reports/api/v2'; // In dev

class TogglScript {
  constructor(togglToken) {
    this.BASIC_AUTH = `${togglToken}:api_token`;
  }

  ///////////////////
  // Basic Methods //
  ///////////////////
  /**
   * GET request on Toggl API
   * https://github.com/toggl/toggl_api_docs/blob/master/toggl_api.md
   * @param {string} path Path for calling the Toggl API, starting with a slash (/)
   * @return {string} response JSON-encoded object; needs to be JSON.parse()ed before using as object
   */
  get(path) {
    let options = {
      'method': 'GET',
      'contentType': 'application/json',
      'headers': { "Authorization": "Basic " + Utilities.base64Encode(this.BASIC_AUTH) },
      'muteHttpExceptions': false
    };
    return UrlFetchApp.fetch(TOGGL_API_VERSION + path, options);
  }
  /**
   * PUT request on Toggl API
   * @param {string} path Path for calling the Toggl API, starting with a slash (/)
   * @param {string} payloadString [Optional] JSON-encoded details of PUT request
   * @return {string} response JSON-encoded object; needs to be JSON.parse()ed before using as object
   */
  put(path, payloadString = null) {
    let options = {
      'method': 'PUT',
      'contentType': 'application/json',
      'headers': { "Authorization": "Basic " + Utilities.base64Encode(this.BASIC_AUTH) },
      'muteHttpExceptions': false
    };
    if (payloadString !== null) {
      options['payload'] = payloadString;
    }
    return UrlFetchApp.fetch(TOGGL_API_VERSION + path, options);
  }
  /**
   * POST request on Toggl API
   * @param {string} path Path for calling the Toggl API, starting with a slash (/)
   * @param {String} payload JSON-encoded details of PUT request
   * @return {String} response JSON-encoded object; needs to be JSON.parse()ed before using as object
   */
  post(path, payload) {
    let options = {
      'method': 'POST',
      'contentType': 'application/json',
      'headers': { "Authorization": "Basic " + Utilities.base64Encode(this.BASIC_AUTH) },
      'muteHttpExceptions': false,
      'payload': payload,
    };
    return UrlFetchApp.fetch(TOGGL_API_VERSION + path, options);
  }
  /**
   * DELETE request on Toggl API
   * @param {string} path Path for calling the Toggl API, starting with a slash (/)
   * @return {string} response HTTP response; '200 OK' when successful
   */
  tDelete(path) {
    let options = {
      'method': 'DELETE',
      'contentType': 'application/json',
      'headers': { "Authorization": "Basic " + Utilities.base64Encode(this.BASIC_AUTH) },
      'muteHttpExceptions': false
    };
    return UrlFetchApp.fetch(TOGGL_API_VERSION + path, options);
  }
  /**
   * Converts a string representing date & time into designated ISO 8601 format.
   * Used when retrieving time entries of a specific period and recording date on spreadsheet.
   * @param {Date} date 
   * @param {string} timeZone 
   * @returns {string} Date string
   */
  togglFormatDate(date, timeZone = 'GMT') {
    return Utilities.formatDate(date, timeZone, "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  /**
   * Converts a string representing date & time into designated ISO 8601 format.
   * Used for updating time entry in Toggl
   * @param {Date} date 
   * @returns {string} Date string
   */
  togglFormatDateUpdate(date) {
    return Utilities.formatDate(date, 'GMT', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  }

  /////////////////////////////////////////////////
  // Methods to Interact with Toggl Time Entries //
  /////////////////////////////////////////////////
  /**
   * GET request for time entries.
   * When startDateString and endDateString is specified, gets the time entries for the period between start and end dates.
   * If not, returns the time entries in the last 9 days. The limit of returned time entries is 1000.
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#get-time-entries-started-in-a-specific-time-range
   * @param {string} startDateString - ISO 8601 date and time strings; e.g., "2020-01-09T19:23:30+09:00"
   * @param {string} endDateString - ISO 8601 date and time strings
   * @return {Object} timeEntries
   */
  getTimeEntries(startDateString = undefined, endDateString = undefined) {
    let extUrl = '';
    if (!startDateString || !endDateString) {
      throw new Error('endDateString is missing');
    } else if (startDateString && endDateString) {
      extUrl = `?start_date=${encodeURIComponent(startDateString)}&end_date=${encodeURIComponent(endDateString)}`;
    }
    let path = '/time_entries' + extUrl;
    return JSON.parse(this.get(path));
  }
  /**
   * Get details of a specific time entry
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#get-time-entry-details
   * @param {number} timeEntryId
   * @return {Object} Details of the time entry
   */
  getTimeEntry(timeEntryId) {
    return JSON.parse(this.get(`/time_entries/${timeEntryId}`));
  }
  /**
   * Get details of a running time entry
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#get-running-time-entry
   * @return {Object} Details of the running time entry. Note that runningTimeEntry.data.duration is a negative value.
   */
  getRunningTimeEntry() {
    return JSON.parse(this.get('/time_entries/current'));
  }
  /**
   * Get all workspaces which belongs to the token owner.
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/workspaces.md#get-workspaces
   * @return {Object}
   */
  getWorkspaces() {
    return JSON.parse(this.get('/workspaces'));
  }
  /**
   * Get details of the specified workspace.
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/workspaces.md#get-single-workspace
   * @param {number} workspaceId
   * @return {Object}
   */
  getWorkspace(workspaceId) {
    return JSON.parse(this.get(`/workspaces/${workspaceId}`));
  }
  /**
   * Get all project(s) in a specified workspace. To get a successful response, the token owner must be workspace admin.
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/workspaces.md#get-workspace-projects
   * @param {number} workspaceId
   * @return {Object}
   */
  getWorkspaceProjects(workspaceId) {
    return JSON.parse(this.get(`/workspaces/${workspaceId}/projects`));
  }
  /**
   * Get details of a specified project.
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/projects.md#get-project-data
   * @param {number} projectId
   * @return {Object}
   */
  getProject(projectId) {
    return JSON.parse(this.get(`/projects/${projectId}`));
  }
  /**
   * POST request to create a time entry
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#create-a-time-entry
   * @param {String} payloadString JSON-encoded update details which should be in form of
   * {'time_entry' : {'wid' : workspaceId, 'pid' : projectId, 'description' : description, 'tags' : tagArray, 'start' : startToggl, 'stop' : stopToggl, 'duration' : duration}
   * @return {Object} createdEntry The created time entry.
   */
  createTimeEntry(payloadString) {
    return JSON.parse(this.post('/time_entries', payloadString));
  }
  /**
   * PUT request to stop a time entry
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#stop-a-time-entry
   * @param {number} timeEntryId ID of the time entry to stop
   * @return {Object} The stopped time entry
   */
  stopTimeEntry(timeEntryId) {
    return JSON.parse(this.put(`/time_entries/${timeEntryId}/stop`));
  }
  /**
   * PUT request to bulk update time entries tags
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#bulk-update-time-entries-tags
   * @param {Array} timeEntryIds Array of time entry IDs to update tags
   * @param {Array} tags Array of tags in string
   * @param {string} tagAction [Optional] String of either 'add' or 'remove'. When this is left blank, update will override all existing tags on the time entries.
   * @return {Object} updatedTimeEntries The updated time entries
   */
  bulkUpdateTags(timeEntryIds, tags, tagAction = null) {
    let payload = { 'time_entry': { 'tags': tags } };
    if (tagAction) {
      payload.time_entry['tag_action'] = tagAction;
    }
    return JSON.parse(this.put(`/time_entries/${encodeURIComponent(timeEntryIds.join())}`, JSON.stringify(payload)));
  }
  /**
   * Stop the running time entry
   * @return {Object} stoppedTimeEntry The stopped time entry
   */
  stopRunningTimeEntry() {
    return this.stopTimeEntry(this.getRunningTimeEntry().data.id);
  }
  /**
   * PUT request to update a time entry
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#update-a-time-entry
   * @param {number} timeEntryId ID of the time entry to update
   * @param {String} payloadString JSON-encoded update details which should be in form of {'time_entry' : {'wid' : workspaceId, 'pid' : projectId, 'description' : description, 'tags' : tagArray, 'start' : startToggl, 'stop' : stopToggl, 'duration' : duration}
   * @return {Object} The updated time entry.
   */
  updateTimeEntry(timeEntryId, payloadString) {
    return JSON.parse(this.put(`/time_entries/${timeEntryId}`, payloadString));
  }
  /**
   * DELETE request to delete a time entry
   * https://github.com/toggl/toggl_api_docs/blob/master/chapters/time_entries.md#delete-a-time-entry
   * @param {number} timeEntryId ID of the time entry to delete
   * @return {string} HTTP response; '200 OK' if successful
   */
  deleteTimeEntry(timeEntryId) {
    return this.tDelete(`/time_entries/${timeEntryId}`);
  }

  //////////////////////////////////////////////////
  // Experimental Methods using Toggl Reports API //
  //////////////////////////////////////////////////
  _getReports(path) {
    let options = {
      'method': 'GET',
      'contentType': 'application/json',
      'headers': { "Authorization": "Basic " + Utilities.base64Encode(this.BASIC_AUTH) },
      'muteHttpExceptions': false
    };
    return UrlFetchApp.fetch(_REPORTS_API_VERSION + path, options);
  }
  /**
   * Get detailed time entries via Reports API
   * https://github.com/toggl/toggl_api_docs/blob/master/reports/detailed.md
   * @param {string} myEmail The name of application or email address so the Toggl team can get in touch in case the API user is doing something wrong [Required by Toggl].
   * @param {number} wid Workspace ID of the workspace from which to retrieve data.
   * @param {string} since ISO 8601 date string in yyyy-MM-dd, e.g., "2020-01-09"
   * @param {number} page [Optional] Page number of the paged response. Defaults to 1.
   * @returns Time entries
   */
  _getReportsDetails(myEmail, wid, since, page = 1) {
    return JSON.parse(this._getReports(`/details/?user_agent=${encodeURIComponent(myEmail)}&workspace_id=${wid}&since=${since}&page=${parseInt(page)}`));
  }
}
