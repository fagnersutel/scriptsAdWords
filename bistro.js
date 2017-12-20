// Copyright 2015, Google Inc. All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
 * @name Account Summary Report
 *
 * @overview The Manager Account Summary Report script generates an at-a-glance
 *     report showing the performance of an entire AdWords Manager Account.
 *     https://developers.google.com/adwords/scripts/docs/solutions/mccapp-account-summary
 *     for more details.
 *
 * @author AdWords Scripts Team [adwords-scripts@googlegroups.com]
 *
 * @version 1.1
 *
 * @changelog
 * - version 1.1
 *   - Add user-updateable fields, and ensure report row ordering.
 * - version 1.0.1
 *   - Added validation for external spreadsheet setup.
 * - version 1.0
 *   - Released initial version.
 */

// The hour of the day at or after which to trigger the process of collating
// the Manager Account Report for yesterday's data. Set at least 3 hours into
// the day to ensure that data for yesterday is complete.
var TRIGGER_NEW_DAY_REPORT_HOUR = 5;
var MILLIS_PER_DAY = 24 * 3600 * 1000;
var MIN_NEW_DAY_REPORT_HOUR = 3;
var MAX_NEW_DAY_REPORT_HOUR = 24;

// The maximum number of accounts within the manager account that can be
// processed in a given day.
var MAX_PARALLEL_ACCOUNTS = 50;
var MAX_ACCOUNTS_IN_MANAGER_ACCOUNT = MAX_PARALLEL_ACCOUNTS * 24;
var MAX_ACCOUNTS_EXCEEDED_ERROR_MSG = 'There are too many accounts within ' +
    'this manager account structure for this script to be used, please ' +
    'consider alternatives for manager account reporting.';

// Take a copy (File -> Make a copy ...) from http://goo.gl/yFkbH6
var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1v9BmO1AYPGyKpQYaH6ksptD50WlUot4xCIsGvDjwJMw/edit?usp=sharing';
var REPORTING_OPTIONS = {
  // Comment out the following line to default to the latest reporting version.
  apiVersion: 'v201705'
};

var DEFAULT_EMPTY_EMAIL = 'foo@example.com';

// The metrics to be pulled back from Account Report.
var QUERY_FIELDS = ['Date', 'Cost', 'Impressions', 'Clicks', 'AveragePosition'];

/**
 * The metrics to be presented in the spreadsheet report. To add additional
 * fields to the report, follow the instructions at the link in the header
 * above.
 */
var DISPLAY_FIELDS =
    ['Cost', 'Avg. CPC', 'CTR', 'Avg. Pos.', 'Impressions', 'Clicks'];

var reportState = null;
var spreadsheetAccess = null;

/**
 * Main entry point for the script.
 */
function main() {
  validateParameters();
  spreadsheetAccess = new SpreadsheetAccess(SPREADSHEET_URL, 'Report');
  // Retrieve a list of dates for which to fetch and create new rows.
  var newDates = spreadsheetAccess.getNextDates();
  // Initialise the object used to keep track of and collate report results on
  // Drive.
  reportState = new ReportState();
  reportState.addDatesToQueue(newDates);

  var nextAccounts = reportState.getNextAccounts();
  if (nextAccounts.length) {
    var dateQueue = reportState.getDateQueue();
    if (dateQueue.length) {
      MccApp.accounts()
          .withIds(nextAccounts)
          .executeInParallel(
              'processAccount', 'processIntermediateResults',
              JSON.stringify(dateQueue));
    }
  } else if (reportState.getCompletedDates().length) {
    processFinalResults();
  }
}

/**
 * @typedef {Object} ReportRow
 * @property {string} Date The date in the format YYYY-MM-DD.
 * @property {number} Cost
 * @property {number} Impressions
 * @property {number} Clicks
 * @property {number} AveragePosition
 */

/**
 * Runs the Report query via AWQL on each individual account. A list of dates
 * required are passed in from the calling manager account process. Each account
 * determines whether it is ready to request each of those dates: A sub account
 * of a manager accountcan have a different timezone to that of the manager
 * account, and therefore it is necessary to check on each account with the local timezone.
 *
 * @param {string} dateQueueJson JSON string representing a list of dates to
 *     process, in ascending date order.
 * @return {string} Stringified Object.<ReportRow>
 */
function processAccount(dateQueueJson) {
  var dateQueue = JSON.parse(dateQueueJson);
  // It is necessary to represent the dates for yesterday and today in local
  // format.
  var tz = AdWordsApp.currentAccount().getTimeZone();
  var today = new Date();
  var yesterday = new Date((new Date()).getTime() - MILLIS_PER_DAY);
  var yesterdayString = Utilities.formatDate(yesterday, tz, 'yyyyMMdd');
  var results = {};
  for (var i = 0; i < dateQueue.length; i++) {
    var nextDate = dateQueue[i];
    // Only retrieve the report if either (a) the date in question is earlier
    // than yesterday, or (b) the date in question is yesterday *and*
    // sufficient hours have passed for yesterday's results to be complete.
    if (nextDate < yesterdayString ||
        (nextDate === yesterdayString &&
         parseInt(Utilities.formatDate(today, tz, 'H')) >=
             TRIGGER_NEW_DAY_REPORT_HOUR)) {
      results[nextDate] = getReportRows(nextDate);
    }
  }
  return JSON.stringify(results);
}

/**
 * Retrieves a row from Account Performance Report for a specified date.
 *
 * @param {string} dateString The date in the form YYYYMMDD.
 * @return {ReportRow}
 */
function getReportRows(dateString) {
  var row = {};
  var report = AdWordsApp.report(
      'SELECT ' + QUERY_FIELDS.join(',') + ' ' +
          'FROM ACCOUNT_PERFORMANCE_REPORT ' +
          'DURING ' + dateString + ',' + dateString,
      REPORTING_OPTIONS);
  if (report.rows().hasNext()) {
    row = report.rows().next();
  } else {
    QUERY_FIELDS.forEach(function(metric) {
      row[metric] = '0';
    });
    row.Date = separateDateString(dateString);
  }
  return row;
}

/**
 * Callback function called on completion of executing managed accounts. Adds
 * all the returned results to the ReportState object and then stores to Drive.
 *
 * @param {Array.<MccApp.ExecutionResult>} executionResultsList
 */
function processIntermediateResults(executionResultsList) {
  reportState = new ReportState();
  for (var i = 0; i < executionResultsList.length; i++) {
    var executionResult = executionResultsList[i];
    var customerId = executionResult.getCustomerId();
    var error = executionResult.getError();
    if (error) {
      Logger.log(
          'Error encountered processing account ' + customerId + ': ' + error);
    } else {
      var results = JSON.parse(executionResult.getReturnValue());
      var completedDates = Object.keys(results);
      for (j = 0; j < completedDates.length; j++) {
        var completedDate = completedDates[j];
        reportState.updateAccountResult(
            customerId, completedDate, results[completedDate]);
      }
    }
  }
  // Save changes to object on Drive.
  reportState.flush();
  if (reportState.getCompletedDates().length) {
    processFinalResults();
  }
}

/**
 * Writes any completed records - where statistics have been returned from all
 * managed accounts and aggregated - to the spreadsheet and optionally sends an
 * email alert.
 */
function processFinalResults() {
  spreadsheetAccess = new SpreadsheetAccess(SPREADSHEET_URL, 'Report');
  var completedResults = reportState.getCompletedDates();
  if (completedResults.length) {
    var isSingleCurrency = reportState.isSingleCurrency();
    for (var i = 0; i < completedResults.length; i++) {
      var rows = completedResults[i].reportData;

      // Step 1: Running totals
      // For each new row, set up variables to store running totals.
      var result = {impressions: 0, clicks: 0, cost: 0, positionSum: 0};
      for (var j = 0; j < rows.length; j++) {
        // Each row of data represents a different account.
        var row = rows[j];
        // Cost, for example, requires only summing Cost across all accounts.
        result.cost += row.Cost;
        result.impressions += row.Impressions;
        result.clicks += row.Clicks;
        // To calculate Average Position, it must be weighted by impressions,
        // then divided by impressions once all accounts have been added.
        result.positionSum += row.Impressions * row.AveragePosition;
      }

      // Step 2: Final aggregation and presentation
      // Perform the final formatting to create a new row.
      var formattedRow = [
        separateDateString(completedResults[i].dateString),
        // Cost is an example where if different sub-accounts have different
        // currencies, adding them together is not meaningful. The below adds
        // "N/A" for "Not Applicable" in this case.
        isSingleCurrency ? result.cost : 'N/A',
        isSingleCurrency ? (result.cost / result.clicks).toFixed(2) : 'N/A',
        // CTR is calculated from dividing total clicks by total impressions,
        // not by summing CTRs from individual accounts.
        (result.clicks * 100 / result.impressions).toFixed(2),
        // Average position is calculated by dividing the sum of all the
        // impression positions, by the number of impressions.
        (result.positionSum / result.impressions).toFixed(1),
        result.impressions, result.clicks
      ];

      spreadsheetAccess.writeNextEntry(formattedRow);
      spreadsheetAccess.sortReportRows();
      spreadsheetAccess.setDateComplete();
      reportState.removeDateFromQueue(completedResults[i].dateString);
    }
    var email = spreadsheetAccess.getEmail();
    if (email) {
      sendEmail(email);
    }
  }
}

/**
 * Constructs and sends email summary.
 *
 * @param {string} email The recipient's email address.
 */
function sendEmail(email) {
  var yesterdayRow = spreadsheetAccess.getPreviousRow(1);
  var twoDaysAgoRow = spreadsheetAccess.getPreviousRow(2);
  var weekAgoRow = spreadsheetAccess.getPreviousRow(5);

  var yesterdayColHeading = yesterdayRow ? yesterdayRow[0] : '-';
  var twoDaysAgoColHeading = twoDaysAgoRow ? twoDaysAgoRow[0] : '-';
  var weekAgoColHeading = weekAgoRow ? weekAgoRow[0] : '-';

  var html = [];
  html.push(
      '<html>', '<body>',
      '<table width=800 cellpadding=0 border=0 cellspacing=0>', '<tr>',
      '<td colspan=2 align=right>',
      '<div style=\'font: italic normal 10pt Times New Roman, serif; ' +
          'margin: 0; color: #666; padding-right: 5px;\'>' +
          'Powered by AdWords Scripts</div>',
      '</td>', '</tr>', '<tr bgcolor=\'#3c78d8\'>', '<td width=500>',
      '<div style=\'font: normal 18pt verdana, sans-serif; ' +
          'padding: 3px 10px; color: white\'>Account Summary report</div>',
      '</td>', '<td align=right>',
      '<div style=\'font: normal 18pt verdana, sans-serif; ' +
          'padding: 3px 10px; color: white\'>',
      AdWordsApp.currentAccount().getCustomerId(), '</h1>', '</td>', '</tr>',
      '</table>', '<table width=800 cellpadding=0 border=0 cellspacing=0>',
      '<tr bgcolor=\'#ddd\'>', '<td></td>',
      '<td style=\'font: 12pt verdana, sans-serif; ' +
          'padding: 5px 0px 5px 5px; background-color: #ddd; ' +
          'text-align: left\'>',
      yesterdayColHeading, '</td>',
      '<td style=\'font: 12pt verdana, sans-serif; ' +
          'padding: 5px 0px 5px 5px; background-color: #ddd; ' +
          'text-align: left\'>',
      twoDaysAgoColHeading, '</td>',
      '<td style=\'font: 12pt verdana, sans-serif; ' +
          'padding: 5px 0px 5x 5px; background-color: #ddd; ' +
          'text-align: left\'>',
      weekAgoColHeading, '</td>', '</tr>');
  for (var d = 0; d < DISPLAY_FIELDS.length; d++) {
    var fieldName = DISPLAY_FIELDS[d];
    html.push(
        emailRow(fieldName, d + 1, yesterdayRow, twoDaysAgoRow, weekAgoRow));
  }
  html.push('</table>', '</body>', '</html>');
  MailApp.sendEmail(
      email, 'AdWords Account ' + AdWordsApp.currentAccount().getCustomerId() +
          ' Summary Report',
      '', {htmlBody: html.join('\n')});
}

/**
 * Constructs a row for embedding in the email message.
 *
 * @param {string} title The title for the row.
 * @param {number} column The index into each ReportRow object for the value to
 *     extract.
 * @param {ReportRow} yesterdayRow Statistics from yesterday, or the most recent
 *     last day processed.
 * @param {ReportRow} twoDaysAgoRow Statistics from 2 days ago, or the 2nd most
 *     recent day processed.
 * @param {ReportRow} weekAgoRow Statistics from a week ago, or the 7th most
 *     recent day processed.
 * @return {string} HTML representing a row of statistics.
 */
function emailRow(title, column, yesterdayRow, twoDaysAgoRow, weekAgoRow) {
  var html = [];
  var twoDaysAgoCell = '<td></td>';
  var weekAgoCell = '<td></td>';
  if (twoDaysAgoRow) {
    twoDaysAgoCell = '<td style=\'padding: 0px 10px\'>' +
        twoDaysAgoRow[column] +
        formatChangeString(yesterdayRow[column], twoDaysAgoRow[column]) +
        '</td>';
  }
  if (weekAgoRow) {
    weekAgoCell = '<td style=\'padding: 0px 10px\'>' + weekAgoRow[column] +
        formatChangeString(yesterdayRow[column], weekAgoRow[column]) + '</td>';
  }
  html.push(
      '<tr>', '<td style=\'padding: 5px 10px\'>' + title + '</td>',
      '<td style=\'padding: 0px 10px\'>' + yesterdayRow[column] + '</td>',
      twoDaysAgoCell, weekAgoCell, '</tr>');
  return html.join('\n');
}

/**
 * Formats HTML representing the change from an old to a new value in the email
 * summary.
 *
 * @param {number} newValue
 * @param {number} oldValue
 * @return {string} HTML representing the change.
 */
function formatChangeString(newValue, oldValue) {
  var newValueString = newValue.toString();
  var oldValueString = oldValue.toString();
  var x = newValueString.indexOf('%');
  if (x != -1) {
    newValueString = newValueString.substring(0, x);
    var y = oldValueString.indexOf('%');
    oldValueString = oldValueString.substring(0, y);
  }

  var change = parseFloat(newValueString - oldValueString).toFixed(2);
  var changeString = change;
  if (x != -1) {
    changeString = change + '%';
  }

  var color = 'cc0000';
  var template = '<span style=\'color: #%s; font-size: 8pt\'> (%s)</span>';
  if (change >= 0) {
    color = '38761d';
  }
  return Utilities.formatString(template, color, changeString);
}

/**
 * Convenience function fo reformat a string date from YYYYMMDD to YYYY-MM-DD.
 *
 * @param {string} date String in form YYYYMMDD.
 * @return {string} String in form YYYY-MM-DD.
 */
function separateDateString(date) {
  return [date.substr(0, 4), date.substr(4, 2), date.substr(6, 2)].join('-');
}

/**
 * @typedef {Object} AccountData
 * @property {string} currencyCode
 * @property {Object.<ReportRow>} records Results for individual dates.
 */

/**
 * @typedef {Object} State
 * @property {Array.<string>} dateQueue Holds an ordered list of dates requiring
 *    report entries.
 * @property {Object.<AccountData>} accounts Holds intermediate results for each
 *    account.
 */

/**
 * ReportState coordinates the ordered retrieval of report data across CIDs, and
 * determines when data is ready for writing to the spreadsheet.
 *
 * @constructor
 */
function ReportState() {
  this.state_ = this.loadOrCreateState_();
}

/**
 * Either loads an existing state representation from Drive, or if one does not
 * exist, creates a new state representation.
 *
 * @return {State}
 * @private_
 */
ReportState.prototype.loadOrCreateState_ = function() {
  var reportStateFiles =
      DriveApp.getRootFolder().getFilesByName(this.getFilename_());
  if (reportStateFiles.hasNext()) {
    var reportStateFile = reportStateFiles.next();
    if (reportStateFiles.hasNext()) {
      this.throwDuplicateFileException_();
    }
    reportState = JSON.parse(reportStateFile.getBlob().getDataAsString());
    this.updateAccountsList_(reportState);
  } else {
    reportState = this.createNewState_();
  }
  return reportState;
};

/**
 * Creates a new state representation on Drive.
 *
 * @return {State}
 * @private
 */
ReportState.prototype.createNewState_ = function() {
  var accountDict = {};
  var accounts = MccApp.accounts().get();
  while (accounts.hasNext()) {
    var account = accounts.next();
    accountDict[account.getCustomerId()] = {
      records: {},
      currencyCode: account.getCurrencyCode()
    };
  }

  var reportState = {dateQueue: [], accounts: accountDict};
  DriveApp.getRootFolder().createFile(
      this.getFilename_(), JSON.stringify(reportState));
  return reportState;
};

/**
 * Updates the state object to reflect both accounts that are added to
 * the manager account and accounts that are removed.
 *
 * @param {State} reportState The state as loaded from Drive.
 * @private_
 */
ReportState.prototype.updateAccountsList_ = function(reportState) {
  var accountState = reportState.accounts;
  var accounts = MccApp.accounts().get();
  var accountDict = {};
  while (accounts.hasNext()) {
    var account = accounts.next();
    var customerId = account.getCustomerId();
    accountDict[customerId] = true;
    if (!accountState.hasOwnProperty(customerId)) {
      accountState[customerId] = {
        records: {},
        currencyCode: account.getCurrencyCode()
      };
    }
  }
  var forRemoval = [];
  var existingAccounts = Object.keys(accountState);
  for (var i = 0; i < existingAccounts.length; i++) {
    if (!accountDict.hasOwnProperty(existingAccounts[i])) {
      forRemoval.push(existingAccounts[i]);
    }
  }
  forRemoval.forEach(function(customerId) { delete accountState[customerId]; });
};

/**
 * Adds dates to the state object, for which reports should be retrieved.
 *
 * @param {Array.<string>} dateList A list of strings in the form YYYYMMDD, that
 *     are to be marked as for report retrieval by each managed account.
 */
ReportState.prototype.addDatesToQueue = function(dateList) {
  if (dateList.length) {
    for (var i = 0; i < dateList.length; i++) {
      var dateString = dateList[i];
      if (this.state_.dateQueue.indexOf(dateString) === -1) {
        this.state_.dateQueue.push(dateString);
      }
    }
    // Ensure the date queue is sorted oldest to newest.
    this.state_.dateQueue.sort();
    this.flush();
  }
};

/**
 * Retrieve the list of dates requiring report generation.
 *
 * @return {Array.<string>} An ordered list of strings in the form YYYYMMDD.
 */
ReportState.prototype.getDateQueue = function() {
  return this.state_.dateQueue;
};

/**
 * Removes a date from the list of dates remaining to have their reports pulled
 * and aggregated, and removes any associated saved statistics from the state
 * object also. Saves the state to Drive.
 *
 * @param {string} dateString Date in the format YYYYMMDD.
 */
ReportState.prototype.removeDateFromQueue = function(dateString) {
  var index = this.state_.dateQueue.indexOf(dateString);
  if (index > -1) {
    this.state_.dateQueue.splice(index, 1);
  }
  var accounts = this.state_.accounts;
  var accountKeys = Object.keys(accounts);
  for (var i = 0; i < accountKeys.length; i++) {
    var customerId = accountKeys[i];
    var records = accounts[customerId].records;
    if (records.hasOwnProperty(dateString)) {
      delete records[dateString];
    }
  }
  this.flush();
};

/**
 * Stores results for a given account in the state object. Does not save to
 * Drive: As this may be called ~50 times in succession for each managed
 * account, call .flush() after all calls to save only once.
 *
 * @param {string} customerId The customerId for the results.
 * @param {string} dateString The date of the results in the form YYYYMMDD.
 * @param {ReportRow} results Statistics from Account Performance Report.
 */
ReportState.prototype.updateAccountResult = function(
    customerId, dateString, results) {
  var accounts = this.state_.accounts;
  if (accounts.hasOwnProperty(customerId)) {
    var records = accounts[customerId].records;
    records[dateString] = results;
  }
};

/**
 * Saves the report state object to Drive.
 */
ReportState.prototype.flush = function() {
  var reportStateFilename = this.getFilename_();
  var reportFiles =
      DriveApp.getRootFolder().getFilesByName(reportStateFilename);
  if (reportFiles.hasNext()) {
    var reportFile = reportFiles.next();
    if (reportFiles.hasNext()) {
      this.throwDuplicateFileException_();
    }
    reportFile.setContent(JSON.stringify(this.state_));
  } else {
    this.throwNoReportFileFoundException_();
  }
};

/**
 * Retrieves the list of accounts to process next. Return accounts in an
 * ordering where those accounts with the oldest incomplete date return first.
 *
 * @return {Array.<string>} A list of CustomerId values.
 */
ReportState.prototype.getNextAccounts = function() {
  var nextAccounts = [];
  var accounts = this.state_.accounts;
  // Sort only to make it easier to test.
  var accountKeys = Object.keys(accounts).sort();
  // dateQueue is ordered from oldest to newest
  var dates = this.state_.dateQueue;
  var i = 0;
  var j = 0;
  while (i < dates.length && nextAccounts.length < MAX_PARALLEL_ACCOUNTS) {
    var date = dates[i];
    while (j < accountKeys.length &&
           nextAccounts.length < MAX_PARALLEL_ACCOUNTS) {
      var customerId = accountKeys[j];
      var records = accounts[customerId].records;
      if (!records.hasOwnProperty(date)) {
        nextAccounts.push(customerId);
      }
      j++;
    }
    i++;
  }
  return nextAccounts;
};

/**
 * @typedef {object} CompletedDate
 * @property {!string} dateString The date of the report data, in YYYYMMDD
 *     format.
 * @property {Array.<ReportRow>} reportData Rows of report data taken from each
 *     account within the manager account.
 */

/**
 * Gets a list of the dates, and associated report data in the State object for
 * which all accounts have data (and are therefore ready for aggregation and
 * writing to a Spreadsheet).
 *
 * @return {!Array.<CompletedDate>} An array of CompletedDate objects, ordered
 *     from the oldest date to the most recent.
 */
ReportState.prototype.getCompletedDates = function() {
  var completedDates = [];
  var dateQueue = this.state_.dateQueue;
  for (var i = 0; i < dateQueue.length; i++) {
    completedDates.push({dateString: dateQueue[i], reportData: []});
  }
  var accounts = this.state_.accounts;
  var accountKeys = Object.keys(accounts);
  for (var j = 0; j < accountKeys.length; j++) {
    var customerId = accountKeys[j];
    var records = accounts[customerId].records;
    var forRemoval = [];
    for (var k = 0; k < completedDates.length; k++) {
      var dateString = completedDates[k].dateString;
      if (records.hasOwnProperty(dateString)) {
        completedDates[k].reportData.push(records[dateString]);
      } else {
        forRemoval.push(k);
      }
    }
    forRemoval.forEach(function(index) { completedDates.splice(index, 1); });
  }
  return completedDates;
};

/**
 * Generate a filename unique to this manager account for saving the
 * intermediate data on Drive.
 *
 * @return {string} The filename.
 * @private
 */
ReportState.prototype.getFilename_ = function() {
  return AdWordsApp.currentAccount().getCustomerId() + '-account-report.json';
};

/**
 * Returns whether the accounts store in the state object all have the same
 * currency or not. This is relevant in determining whether showing an
 * aggregated cost and CTR is meaningful.
 *
 * @return {boolean} True if only one currency is present.
 */
ReportState.prototype.isSingleCurrency = function() {
  var accounts = this.state_.accounts;
  var accountKeys = Object.keys(accounts);
  for (var i = 1; i < accountKeys.length; i++) {
    if (accounts[accountKeys[i - 1]].currencyCode !==
        accounts[accountKeys[i]].currencyCode) {
      return false;
    }
  }
  return true;
};

/**
 * Sets the currency code for a given account.
 *
 * @param {string} customerId
 * @param {string} currencyCode , e.g. 'USD'
 */
ReportState.prototype.setCurrencyCode = function(customerId, currencyCode) {
  var accounts = this.state_.accounts;
  if (accounts.hasOwnProperty(customerId)) {
    accounts[customerId].currencyCode = currencyCode;
  }
};

/**
 * Throws an exception if there are multiple files with the same name.
 *
 * @private
 */
ReportState.prototype.throwDuplicateFileException_ = function() {
  throw 'Multiple files named ' + this.getFileName_() + ' detected. Please ' +
      'ensure there is only one file named ' + this.getFileName_() +
      ' and try again.';
};

/**
 * Throws an exception for when no file is found for the given name.
 *
 * @private
 */
ReportState.prototype.throwNoReportFileFoundException_ = function() {
  throw 'Could not find the file named ' + this.getFileName_() +
      ' to save the to.';
};

/**
 * Class used to ease reading and writing to report spreadsheet.
 *
 * @param {string} spreadsheetUrl
 * @param {string} sheetName The sheet name to read/write results from/to.
 * @constructor
 */
function SpreadsheetAccess(spreadsheetUrl, sheetName) {
  // Offsets into the existing template sheet for the top left of the data.
  this.DATA_COL_ = 2;
  this.DATA_ROW_ = 6;
  this.spreadsheet_ = validateAndGetSpreadsheet(spreadsheetUrl);
  this.sheet_ = this.spreadsheet_.getSheetByName(sheetName);
  this.accountTz_ = AdWordsApp.currentAccount().getTimeZone();
  this.spreadsheetTz_ = this.spreadsheet_.getSpreadsheetTimeZone();
  this.spreadsheet_.getRangeByName('account_id_report')
      .setValue(AdWordsApp.currentAccount().getCustomerId());

  var d = new Date();
  d.setSeconds(0);
  d.setMilliseconds(0);

  var s = new Date(
      Utilities.formatDate(d, this.spreadsheetTz_, 'MMM dd,yyyy HH:mm:ss'));
  this.spreadsheetOffset_ = s.getTime() - d.getTime();
}

/**
 * Transforms a Date object as read from the spreadsheet into a Date object
 * which can be used to obtain the same Year, Month, Day, Hours values as would
 * be displayed in the spreadsheet.
 *
 * @param {Date} date
 * @return {Date} A date object shifted by the difference in timezones.
 * @private
 */
SpreadsheetAccess.prototype.localDateToSpreadsheetDate_ = function(date) {
  var spreadsheetSecs = date.getTime() - this.spreadsheetOffset_;
  return new Date(spreadsheetSecs);
};

/**
 * Retrieves a list of dates for which Account Report data is required. This is
 * based on the last entry in the spreadsheet. If the last entry value is empty
 * then yesterday is used, otherwise, all dates between the last entry and
 * yesterday are used, except those for which data is already in the Sheet.
 *
 * @return {Array.<string>} List of dates in YYYYMMDD format.
 */
SpreadsheetAccess.prototype.getNextDates = function() {
  var nextDates = [];
  var y = new Date((new Date()).getTime() - MILLIS_PER_DAY);
  var yesterday = Utilities.formatDate(y, this.accountTz_, 'yyyyMMdd');
  var lastCheck = this.spreadsheet_.getRangeByName('last_check').getValue();

  if (lastCheck.length === 0) {
    nextDates = [yesterday];
  } else {
    var lastCheckDate =
        Utilities.formatDate(lastCheck, this.spreadsheetTz_, 'yyyyMMdd');
    while (lastCheckDate !== yesterday) {
      lastCheck.setTime(lastCheck.getTime() + MILLIS_PER_DAY);
      lastCheckDate =
          Utilities.formatDate(lastCheck, this.spreadsheetTz_, 'yyyyMMdd');
      nextDates.push(lastCheckDate);
    }
  }

  var sheet = this.spreadsheet_.getSheetByName('Report');
  var data = sheet.getDataRange().getValues();
  var existingDates = {};
  data.slice(5).forEach(function(row) {
    var existingDate =
        Utilities.formatDate(row[1], this.spreadsheetTz_, 'yyyyMMdd');
    existingDates[existingDate] = true;
  });
  return nextDates.filter(function(d) {
    return !existingDates[d];
  });
};

/**
 * Updates the spreadsheet to set the date for the last saved report data.
 */
SpreadsheetAccess.prototype.setDateComplete = function() {
  var sheet = this.spreadsheet_.getSheetByName('Report');
  var data = sheet.getDataRange().getValues();
  if (data.length > 5) {
    var lastDate = data[data.length - 1][1];
    this.spreadsheet_.getRangeByName('last_check').setValue(lastDate);
  }
};

/**
 * Writes the next row of report data to the spreadsheet.
 *
 * @param {Array.<*>} row An array of report values
 */
SpreadsheetAccess.prototype.writeNextEntry = function(row) {
  var lastRow = this.sheet_.getDataRange().getLastRow();
  if (lastRow + 1 > this.sheet_.getMaxRows()) {
    this.sheet_.insertRowAfter(lastRow);
  }
  this.sheet_.getRange(lastRow + 1, this.DATA_COL_, 1, row.length).setValues([
    row
  ]);
};

/**
 * Retrieves the values for a previously written row
 *
 * @param {number} daysAgo The reversed index of the required row, e.g. 1 is the
 *     last written row, 2 is the one before that etc.
 * @return {Array.<*>} The array data, or null if the index goes out of bounds.
 */
SpreadsheetAccess.prototype.getPreviousRow = function(daysAgo) {
  var index = this.sheet_.getDataRange().getLastRow() - daysAgo + 1;
  if (index < this.DATA_ROW_) {
    return null;
  }
  var numColumns = DISPLAY_FIELDS.length;
  var row = this.sheet_.getRange(index, this.DATA_COL_, 1, numColumns + 1)
                .getValues()[0];
  row[0] = Utilities.formatDate(row[0], this.spreadsheetTz_, 'yyyy-MM-dd');
  return row;
};

/**
 * Retrieves the email address set in the spreadsheet.
 *
 * @return {string}
 */
SpreadsheetAccess.prototype.getEmail = function() {
  return this.spreadsheet_.getRangeByName('email').getValue();
};

/**
 * Sorts the data in the spreadsheet into ascending date order.
 */
SpreadsheetAccess.prototype.sortReportRows = function() {
  var sheet = this.spreadsheet_.getSheetByName('Report');

  var data = sheet.getDataRange().getValues();
  var reportRows = data.slice(5);
  if (reportRows.length) {
    reportRows.sort(function(rowA, rowB) {
      if (!rowA || !rowA.length) {
        return -1;
      } else if (!rowB || !rowB.length) {
        return 1;
      } else if (rowA[1] < rowB[1]) {
        return -1;
      } else if (rowA[1] > rowB[1]) {
        return 1;
      }
      return 0;
    });
    sheet.getRange(6, 1, reportRows.length, reportRows[0].length)
        .setValues(reportRows);
  }
};

/**
 * Validates the parameters related to the data retrieval to make sure
 * they are within valid values.
 * @throws {Error} If the new day trigger hour is less than 3 or
 * greater than or equal to 24
 */
function validateParameters() {
  if (TRIGGER_NEW_DAY_REPORT_HOUR < MIN_NEW_DAY_REPORT_HOUR ||
          TRIGGER_NEW_DAY_REPORT_HOUR >= MAX_NEW_DAY_REPORT_HOUR) {
    throw new Error('Please set the new day trigger hour at least 3 hours' +
      ' into the day and less than 24 hours after the start of the day');
  }
}

/**
 * Validates the provided spreadsheet URL and email address
 * to make sure that they're set up properly. Throws a descriptive error message
 * if validation fails.
 *
 * @param {string} spreadsheeturl The URL of the spreadsheet to open.
 * @return {Spreadsheet} The spreadsheet object itself, fetched from the URL.
 * @throws {Error} If the spreadsheet URL or email hasn't been set
 */
function validateAndGetSpreadsheet(spreadsheeturl) {
  if (spreadsheeturl == 'INSERT_SPREADSHEET_URL_HERE') {
    throw new Error('Please specify a valid Spreadsheet URL. You can find' +
        ' a link to a template in the associated guide for this script.');
  }
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheeturl);
  var email = spreadsheet.getRangeByName('email').getValue();
  if (email == DEFAULT_EMPTY_EMAIL) {
    throw new Error('Please either set a custom email address in the' +
        ' spreadsheet, or set the email field in the spreadsheet to blank' +
        ' to send no email.');
  }
  return spreadsheet;
}