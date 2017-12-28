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
 * @name Kratu
 *
 * @overview The Kratu script is a flexible MCC-level report showing several
 *     performance signals for each account visually as a heat map. See
 *     https://developers.google.com/adwords/scripts/docs/solutions/kratu
 *     for more details.
 *
 * @author AdWords Scripts Team [adwords-scripts@googlegroups.com]
 *
 * @version 1.0.2
 *
 * @changelog
 * - version 1.0.2
 *   - Fixed bug with run frequency to allow the script to run daily.
 * - version 1.0.1
 *   - Added validation for external spreadsheet setup.
 *   - Updated reporting version to v201609.
 * - version 1.0
 *   - Released initial version.
 */

var CONFIG = {
  // URL to the main / template spreadsheet
  SPREADSHEET_URL: 'YOUR_SPREADSHEET_URL'
};

/**
 * Configuration to be used for running reports.
 */
var REPORTING_OPTIONS = {
  // Comment out the following line to default to the latest reporting version.
  //apiVersion: 'v201705'
};

/**
 * Main method, coordinate and trigger either new report creation or continue
 * unfinished report.
 */
function main() {
  init();

  if (spreadsheetManager.hasUnfinishedRun()) {
    continueUnfinishedReport();
  } else {
    var reportFrequency = settingsManager.getSetting('ReportFrequency', true);
    var lastReportStart = spreadsheetManager.getLastReportStartTimestamp();

    if (!lastReportStart ||
        dayDifference(lastReportStart, getTimestamp()) >= reportFrequency) {
      startNewReport();
    } else {
      debug('Nothing to do');
    }
  }
}

/**
 * Initialization procedures to be done before anything else.
 */
function init() {
  spreadsheetManager.readSignalDefinitions();
  settingsManager.readSettings();
}

/**
 * Continues an unfinished report. This happens whenever there are accounts
 * that are not processed within the last report. This method picks these
 * up, processes them and marks the report as completed if no accounts are
 * left.
 */
function continueUnfinishedReport() {
  debug('Continuing unfinished report: ' +
    spreadsheetManager.getCurrentRunSheet().getUrl());

  var iterator = spreadsheetManager.getUnprocessedAccountIterator();
  var processed = 0;
  while (iterator.hasNext() &&
    processed++ < settingsManager.getSetting('NumAccountsProcess', true)) {

    var account = iterator.next();
    processAccount(account);
  }

  writeAccountDataToSpreadsheet();

  if (processed > 0 && spreadsheetManager.allAccountsProcessed()) {
    debug('All accounts processed, marking report as complete');

    // Remove protection from sheets, allow changes again
    spreadsheetManager.removeProtection();

    spreadsheetManager.markRunAsProcessed();
    sendEmail();
  }

  debug('Processed ' + processed + ' accounts');
}

/**
 * Creates a new report by copying the report template to a new spreadsheet,
 * gathering all accounts under the MCC and mark them as not processed.
 * Please note that this method will not actually process any accounts.
 */
function startNewReport() {
  debug('Creating new report');

  // Protect the sheets that shouldn't be changed during execution
  spreadsheetManager.setProtection();

  // Delete all account info
  spreadsheetManager.clearAllAccountInfo();

  // Iterate over accounts
  var accountSelector = MccApp.accounts();
  var accountLabel = settingsManager.getSetting('AccountLabel', false);
  if (accountLabel) {
    accountSelector.withCondition("LabelNames CONTAINS '" + accountLabel + "'");
  }
  var accountIterator = accountSelector.get();

  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    debug('Adding account: ' + account.getCustomerId());

    spreadsheetManager.addAccount(account.getCustomerId());
  }

  // Now add the run
  var newRunSheet = spreadsheetManager.addRun();
  debug('New report created at ' + newRunSheet.getUrl());
}

/**
 * Processes a single account.
 *
 * @param {object} account the AdWords account object
 */
function processAccount(account) {
  debug('- Processing ' + account.getCustomerId());
  MccApp.select(account);
  signalManager.processAccount(account);

  spreadsheetManager.markAccountAsProcessed(account.getCustomerId());
}

/**
 * After processing & gathering data for all accounts,
 * write it to the spreadsheet.
 */
function writeAccountDataToSpreadsheet() {
  var accountInfos = signalManager.getAccountInfos();

  spreadsheetManager.writeHeaderRow();

  for (var i = 0; i < accountInfos.length; i++) {
    var accountInfo = accountInfos[i];
    spreadsheetManager.writeDataRow(accountInfo);
  }
}

/**
 * Sends email if an email was provided in the settings.
 * Otherwise does nothing.
 */
var sendEmail = function() {
  var recipientEmail = settingsManager.getSetting('RecipientEmail', false);

  if (recipientEmail) {
    MailApp.sendEmail(recipientEmail,
      'Kratu Report is ready',
      spreadsheetManager.getCurrentRunSheet().getUrl());
    debug('Email sent to ' + recipientEmail);
  }
};

/**
 * Returns the number of days between two timestamps.
 *
 * @param {number} time1 the newer (more recent) timestamps
 * @param {number} time2 the older timestamps
 * @return {number} number of full days between the given dates
 */
var dayDifference = function(time1, time2) {
  return parseInt((time2 - time1) / (24 * 3600 * 1000));
};

/**
 * Returns the current timestamp.
 *
 * @return {number} the current timestamp
 */
function getTimestamp() {
  return new Date().getTime();
}

/**
 * Module for calculating account signals and infos to be shown in the report.
 *
 * @return {object} callable functions corresponding to the available
 * actions
 */
var signalManager = (function() {
  var accountInfos = new Array();

  /**
   * Processes one account, which in 2 steps adds an accountInfo object
   * to the list.
   * - Calculate the raw signals
   * - Postprocess the raw signals (normalize scores, ...)
   *
   * @param {object} account the AdWords account object
   */
  var processAccount = function(account) {
    var rawSignals = calculateRawSignals(account);

    var accountInfo = {
      account: account,
      rawSignals: rawSignals
    };

    processSignals(accountInfo);

    accountInfos.push(accountInfo);
  };

  /**
   * Returns an array of all processed accounts so far. These are ordered by
   * decreasing score.
   *
   * @return {object} array of the accountInfo objects
   */
  var getAccountInfos = function() {
    accountInfos.sort(function(a, b) {
      return b.score - a.score;
    });

    return accountInfos;
  };

  /**
   * Normalizes a raw signal value based in the signal's definition
   * (min, max values).
   *
   * @param {object} signalDefinition definition of the signal
   * @param {number} value numeric value of that signal
   * @return {number} the normalized value
   */
  var normalize = function(signalDefinition, value) {
    var min = signalDefinition.min;
    var max = signalDefinition.max;

    if (signalDefinition.direction == 'High') {
      if (value >= max)
        return 1;
      if (value <= min)
        return 0;

      return (value - min) / (max - min);
    } else if (signalDefinition.direction == 'Low') {
      if (value >= max)
        return 0;
      if (value <= min)
        return 1;

      return 1 - ((value - min) / (max - min));
    } else {
      return value;
    }
  };

  /**
   * Post-processes the raw signals.
   *
   * @param {object} accountInfo the object storing all info about that account
   *                 (including raw signals)
   */
  var processSignals = function(accountInfo) {
    var signalDefinitions = spreadsheetManager.getSignalDefinitions();
    var sumWeights = spreadsheetManager.getSumWeights();
    var sumScore = 0;

    accountInfo.signals = {};

    for (var i = 0; i < signalDefinitions.length; i++) {
      var signalDefinition = signalDefinitions[i];
      if (signalDefinition.includeInReport == 'Yes') {
        var value = accountInfo.rawSignals[signalDefinition.name];

        accountInfo.signals[signalDefinition.name] = {
          definition: signalDefinition,
          value: value,
          displayValue: value
        };

        if (signalDefinition.type == 'Number') {
          var normalizedValue = normalize(signalDefinition, value);
          var signalScore = normalizedValue * signalDefinition.weight;
          sumScore += signalScore;

          accountInfo.signals[signalDefinition.name].normalizedValue =
            normalizedValue;
          accountInfo.signals[signalDefinition.name].signalScore = signalScore;
        }
      }
    }

    accountInfo.scoreSum = sumScore;
    accountInfo.scoreWeights = sumWeights;
    accountInfo.score = sumScore / sumWeights;
  };

  /**
   * Calculate the raw signals.
   *
   * @param {object} account the AdWords account object
   * @return {object} an associative array containing raw signals
   *                  (as name -> value pairs)
   */
  var calculateRawSignals = function(account) {
    // Use reports for signal creation, dynamically create an AWQL query here
    var signalDefinitions = spreadsheetManager.getSignalDefinitions();

    var signalFields = [];
    for (var i = 0; i < signalDefinitions.length; i++) {
      var signalDefinition = signalDefinitions[i];
      signalFields.push(signalDefinition.name);
    }

    var query = 'SELECT ' + signalFields.join(',') +
                ' FROM ACCOUNT_PERFORMANCE_REPORT DURING ' +
                settingsManager.getSetting('ReportPeriod', true);

    var report = AdWordsApp.report(query, REPORTING_OPTIONS);
    var rows = report.rows();

    // analyze the rows (should be only one)
    var rawSignals = {};
    while (rows.hasNext()) {
      var row = rows.next();

      for (var i = 0; i < signalDefinitions.length; i++) {
        var signalDefinition = signalDefinitions[i];

        var value = row[signalDefinition.name];
        if (value.indexOf('%') > -1) {
          value = parseFloat(value) / 100.0;
        }

        rawSignals[signalDefinition.name] = value;
      }

    }

    return rawSignals;
  };

  // Return the external interface.
  return {
    processAccount: processAccount,
    getAccountInfos: getAccountInfos
  };

})();

/**
 * Module for interacting with the spreadhsheets. Offers several
 * functions that other modules can use when storing / retrieving data
 * In general, there are two spreadsheets involved:
 * - a main spreadsheet containing processing information, settings
 *   and a template for the reports
 * - a report spreadsheet for each run (one loop over all accounts)
 *
 * @return {object} callable functions corresponding to the available
 * actions
 */
var spreadsheetManager = (function() {
  validateConfig();
  var spreadsheet = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
  var currentRunSheet = null;
  var accountsTab = spreadsheet.getSheetByName('Accounts');
  var historyTab = spreadsheet.getSheetByName('History');
  var signalsTab = spreadsheet.getSheetByName('Signals');
  var settingsTab = spreadsheet.getSheetByName('Settings');
  var templateTab = spreadsheet.getSheetByName('Template');
  var processedAccounts = 0;
  var signalDefinitions;
  var sumWeights;

  /**
   * Adds protection and notes to all sheets that should not be
   * changed while a report is being processed.
   */
  var setProtection = function() {
    setSheetProtection(signalsTab);
    setSheetProtection(settingsTab);
    setSheetProtection(templateTab);
  };

  /**
   * Adds protection and notes to a sheet / tab.
   *
   * @param {object} the sheet to add protection to
   */
  var setSheetProtection = function(tab) {
    var protection = tab.protect().setDescription(tab.getName() +
                       ' Protection');

    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
    tab.getRange('A1').setNote('A report is currently being executed, ' +
                       'you can not edit this sheet until it is finished.');
  };

  /**
   * Adds a protection and notes to all sheets that should not be
   * changed while a report is being processed.
   */
  var removeProtection = function() {
    removeSheetProtection(signalsTab);
    removeSheetProtection(settingsTab);
    removeSheetProtection(templateTab);
  };

  /**
   * Remove the protection from a sheet / tab.
   *
   * @param {object} the sheet to remove protection from
   */
  var removeSheetProtection = function(tab) {
    var protection = tab.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
    if (protection && protection.canEdit()) {
      protection.remove();
    }
    tab.clearNotes();
  };

  /**
   * Reads and returns the range of settings in the main spreadsheet.
   *
   * @return {object} the range object containing all settings
   */
  var readSettingRange = function() {
    return settingsTab.getRange(2, 1, settingsTab.getLastRow(), 3);
  };

  /**
   * Read and return the signal definitions as defined in the Signals tab
   * of the general spreadsheet. See below for how a signal definition object
   * looks like.
   *
   * @param {object} range the range of cells
   * @return {object} an array of signal definition objects
   */
  var readSignalDefinitions = function() {
    signalDefinitions = new Array();

    var range = signalsTab.getRange(2, 1, signalsTab.getLastRow(), 9);
    var values = range.getValues();
    for (var i = 0; i < range.getNumRows(); i++) {
      if (values[i][0] == '')
        continue;

      var signalDefinition = {
        name: values[i][0],
        displayName: values[i][1],
        includeInReport: values[i][2],
        type: values[i][3],
        direction: values[i][4],
        format: values[i][5],
        weight: values[i][6],
        min: values[i][7],
        max: values[i][8]
      };

      signalDefinitions.push(signalDefinition);
    }

    calculateSumWeights();

    debug('Using ' + signalDefinitions.length + ' signals');
  };

  /**
   * Returns an array of signal definitions to work with.
   *
   * @return {object} array of signal definitions to work with
   */
  var getSignalDefinitions = function() {
    return signalDefinitions;
  };

  /**
   * Returns the sum of weights of all signal definitions
   *
   * @return {number} sum of weights of all signal definitions
   */
  var getSumWeights = function() {
    return sumWeights;
  };

  /**
   * Calculates the overall sum of score weights for normalization of the score.
   */
  var calculateSumWeights = function() {
    sumWeights = 0;

    for (var i = 0; i < signalDefinitions.length; i++) {
      var signalDefinition = signalDefinitions[i];
      if (signalDefinition.type == 'Number' &&
          signalDefinition.includeInReport == 'Yes') {
       sumWeights += signalDefinition.weight;
      }
    }
  };

  /**
   * Adds a "run" (loop over all accounts) to the general spreadsheet.
   */
  var addRun = function() {
    // use formatted date in spreadsheet name and date cell
    var timezone = AdWordsApp.currentAccount().getTimeZone();
    var formattedDate = Utilities.formatDate(new Date(),
                          timezone, 'MMM dd, yyyy');

    var runSpreadsheet = spreadsheet.copy(spreadsheet.getName() +
                          ' - ' + formattedDate);

    runSpreadsheet.deleteSheet(runSpreadsheet.getSheetByName('Accounts'));
    runSpreadsheet.deleteSheet(runSpreadsheet.getSheetByName('History'));
    runSpreadsheet.deleteSheet(runSpreadsheet.getSheetByName('Settings'));
    runSpreadsheet.deleteSheet(runSpreadsheet.getSheetByName('Parameters'));
    runSpreadsheet.deleteSheet(runSpreadsheet.getSheetByName('Signals'));
    runSpreadsheet.getSheetByName('Template').setName('Report');
    removeSheetProtection(runSpreadsheet.getSheetByName('Report'));

    historyTab.appendRow([getTimestamp(), null, runSpreadsheet.getUrl()]);
    historyTab.getRange(historyTab.getLastRow(), 1, 1, 3).clearFormat();

    runSpreadsheet.getRangeByName('AccountID').setValue(
      AdWordsApp.currentAccount().getCustomerId());
    runSpreadsheet.getRangeByName('Date').setValue(formattedDate);

    return runSpreadsheet;
  };

  /**
   * Checks if there is an unfinished (=not all accounts processed yet)
   * report in the run history list.
   *
   * @return {boolean} whether there is an unfinished report
   */
  var hasUnfinishedRun = function() {
    var lastRow = historyTab.getLastRow();

    // has no run at all
    if (lastRow == 1) {
      return false;
    }

    var lastRunEndDate = historyTab.getRange(lastRow, 2, 1, 1).getValue();
    if (lastRunEndDate) {
      return false;
    }

    return true;
  };

  /**
   * Marks the current report (a.k.a run) as finished by adding an end date.
   */
  var markRunAsProcessed = function() {
    var lastRow = historyTab.getLastRow();
    if (lastRow > 1) {
      historyTab.getRange(lastRow, 2, 1, 1).setValue(getTimestamp());
    }
  };

  /**
   * Returns the start timestamp of the last unfinished report.
   *
   * @return {number} the timestamp of the last unfinished report (null if
   *                  there is none)
   */
  var getLastReportStartTimestamp = function() {
    var lastRow = historyTab.getLastRow();
    if (lastRow > 1) {
      return historyTab.getRange(lastRow, 1, 1, 1).getValue();
    } else {
      return null;
    }
  };

  /**
   * Returns the current run sheet to be used for report generation.
   * This is always the last one in the History tab of the general sheet.
   *
   * @return {object} the current run sheet
   */
  var getCurrentRunSheet = function() {
    if (currentRunSheet != null)
      return currentRunSheet;

    var range = historyTab.getRange(historyTab.getLastRow(), 3, 1, 1);
    var url = range.getValue();
    currentRunSheet = SpreadsheetApp.openByUrl(url);
    return currentRunSheet;
  };

  /**
   * Adds an account to the list of 'known' accounts.
   *
   * @param {string} cid the cid of the account
   */
  var addAccount = function(cid) {
    var maxRow = accountsTab.appendRow([cid]);
    accountsTab.getRange(accountsTab.getLastRow(), 1, 1, 2).clearFormat();
  };

  /**
   * Marks an account as processed in the general sheet. Like this,
   * the script can be executed several times and will always
   * run for a batch of unprocessed accounts.
   *
   * @param {string} cid the customer id of the account that has been processed
   */
  var markAccountAsProcessed = function(cid) {
    var range = accountsTab.getRange(2, 1, accountsTab.getLastRow() - 1, 2);

    var values = range.getValues();
    for (var i = 0; i < range.getNumRows(); i++) {
      var rowCid = values[i][0];
      if (cid == rowCid) {
        accountsTab.getRange(i + 2, 2).setValue(getTimestamp());
        processedAccounts++;
      }
    }

  };

  /**
   * Clears the list of 'known' accounts.
   */
  var clearAllAccountInfo = function() {
    var lastRow = accountsTab.getLastRow();

    if (lastRow > 1) {
      accountsTab.deleteRows(2, lastRow - 1);
    }
  };

  /**
   * Creates a selector for the next batch of accounts that are not
   * processed yet.
   *
   * @return {object} a selector that can be used for parallel processing or
   *                  getting an iterator
   */
  var getUnprocessedAccountIterator = function() {
    var accounts = getUnprocessedAccounts();

    var selector = MccApp.accounts().withIds(accounts);
    var iterator = selector.get();
    return iterator;
  };

  /**
   * Reads and returns the next batch of unprocessed accounts from the general
   * spreadsheet.
   *
   * @return {object} an array of unprocessed cids
   */
  var getUnprocessedAccounts = function() {
    var accounts = [];

    var range = accountsTab.getRange(2, 1, accountsTab.getLastRow() - 1, 2);

    for (var i = 0; i < range.getNumRows(); i++) {
      var cid = range.getValues()[i][0];
      var processed = range.getValues()[i][1];

      if (processed != '' || accounts.length >=
              settingsManager.getSetting('NumAccountsProcess', true)) {
        continue;
      }

      accounts.push(cid);
    }

    return accounts;
  };

  /**
   * Scans the list of accounts and returns true if all of them
   * are processed.
   *
   * @return {boolean} true, if all accounts are processed
   */
  var allAccountsProcessed = function() {
    var range = accountsTab.getRange(2, 1, accountsTab.getLastRow() - 1, 2);

    for (var i = 0; i < range.getNumRows(); i++) {
      var cid = range.getValues()[i][0];
      var processed = range.getValues()[i][1];

      if (processed) {
        continue;
      }

      return false;
    }

    return true;
  };

  /**
   * Writes the data headers (signal names) in the current run sheet.
   */
  var writeHeaderRow = function() {
    var sheet = getCurrentRunSheet();
    var reportTab = sheet.getSheetByName('Report');

    var row = [''];
    for (var i = 0; i < signalDefinitions.length; i++) {
      var signalDefinition = signalDefinitions[i];
      if (signalDefinition.includeInReport == 'Yes') {
        row.push(signalDefinition.displayName);
      }
    }
    row.push('Score');

    var range = reportTab.getRange(4, 1, 1, row.length);
    range.setValues([row]);
    range.clearFormat();
    range.setFontWeight('bold');
    range.setBackground('#38c');
    range.setFontColor('#fff');
  };

  /**
   * Writes a row of data (signal values) in the current run sheet.
   *
   * @param {object} accountInfo the accountInfo object containing the
   *                 calculated signals
   */
  var writeDataRow = function(accountInfo) {
    // prepare the data
    var sheet = getCurrentRunSheet();
    var tab = sheet.getSheetByName('Report');

    var row = [''];
    for (var i = 0; i < signalDefinitions.length; i++) {
      var signalDefinition = signalDefinitions[i];
      if (signalDefinition.includeInReport == 'Yes') {
        var displayValue =
             accountInfo.signals[signalDefinition.name].displayValue;

        row.push(displayValue);
      }
    }
    row.push(accountInfo.score);

    // write it
    tab.appendRow(row);

    // now do the formatting
    var currentRow = tab.getLastRow();
    var rowRange = tab.getRange(currentRow, 1, 1, row.length);
    rowRange.clearFormat();

    // arrays for number formats and colors, first fill them with values
    // and later apply to the row
    var dataRange = tab.getRange(currentRow, 2, 1, row.length - 1);
    var fontColors = [[]];
    var backgroundColors = [[]];
    var numberFormats = [[]];
    var colIndex = 0;

    for (var i = 0; i < signalDefinitions.length; i++) {
      var signalDefinition = signalDefinitions[i];
      if (signalDefinition.includeInReport == 'Yes') {
        var value = accountInfo.signals[signalDefinition.name].value;
        var displayValue =
              accountInfo.signals[signalDefinition.name].displayValue;
        var normalizedValue =
              accountInfo.signals[signalDefinition.name].normalizedValue;

        var colors = [2];
        if (signalDefinition.type == 'Number') {
          numberFormats[0][colIndex] = signalDefinition.format;
          colors = getNumberColors(normalizedValue);
        } else if (signalDefinition.type == 'String') {
          colors = getStringColors(value);
        }

        fontColors[0][colIndex] = colors[0];
        backgroundColors[0][colIndex] = colors[1];

        colIndex++;
      }
    }

    // formatting for the score (last column)
    numberFormats[0][colIndex] = '0.00%';
    var scoreColors = getNumberColors(accountInfo.score);
    fontColors[0][colIndex] = scoreColors[0];
    backgroundColors[0][colIndex] = scoreColors[1];

    // now actually apply the formats
    dataRange.setNumberFormats(numberFormats);
    dataRange.setFontColors(fontColors);
    dataRange.setBackgroundColors(backgroundColors);
  };

  /**
   * Helper method for creating the array of colors based on the given
   * setting names.
   *
   * @param {string} settingFontColor name of the setting to use as font color
   * @param {string} settingBackgroundColor name of the setting to use as
   *                                        background color
   * return {object} an array with the colors to apply
   *                 (index 0 -> font color, index 1 -> background color)
   */
  var getColors = function(settingFontColor, settingBackgroundColor) {
     var colors = [];

     colors[0] = settingsManager.getSetting(settingFontColor, false);
     colors[1] = settingsManager.getSetting(settingBackgroundColor, false);

     return colors;
  };

  /**
   * Helper method for returning the "string" colors for a certain value.
   *
   * @param {string} stringValue the value of the cell
   * return {object} an array with the colors to apply
   *                 (index 0 -> font color, index 1 -> background color)
   */
  var getStringColors = function(stringValue) {
     return getColors('StringFgColor', 'StringBgColor');
  };

  /**
   * Helper method for applying the "number" format to a certain range.
   * Numeric value cells have different formats depending on their score value
   * (defined by the settings), this method applies these formats.
   *
   * @param {number} numericValue the value of the cell
   * return {object} an array with the colors to apply
   *                 (index 0 -> font color, index 1 -> background color)
   */
  var getNumberColors = function(numericValue) {
    var level1MinValue = settingsManager.getSetting('Level1MinValue', false);
    var level2MinValue = settingsManager.getSetting('Level2MinValue', false);
    var level3MinValue = settingsManager.getSetting('Level3MinValue', false);
    var level4MinValue = settingsManager.getSetting('Level4MinValue', false);
    var level5MinValue = settingsManager.getSetting('Level5MinValue', false);

    if (level5MinValue && numericValue > level5MinValue) {
      return getColors('Level5FgColor', 'Level5BgColor');
    } else if (level4MinValue && numericValue > level4MinValue) {
      return getColors('Level4FgColor', 'Level4BgColor');
    } else if (level3MinValue && numericValue > level3MinValue) {
      return getColors('Level3FgColor', 'Level3BgColor');
    } else if (level2MinValue && numericValue > level2MinValue) {
      return getColors('Level2FgColor', 'Level2BgColor');
    } else if (level1MinValue && numericValue > level1MinValue) {
      return getColors('Level1FgColor', 'Level1BgColor');
    }

    // if no level reached, no coloring
    var defaultColors = [null, null];
    return defaultColors;
  };

  // Return the external interface.
  return {
    setProtection: setProtection,
    removeProtection: removeProtection,
    readSettingRange: readSettingRange,
    readSignalDefinitions: readSignalDefinitions,
    getSignalDefinitions: getSignalDefinitions,
    getSumWeights: getSumWeights,
    addRun: addRun,
    hasUnfinishedRun: hasUnfinishedRun,
    markRunAsProcessed: markRunAsProcessed,
    getLastReportStartTimestamp: getLastReportStartTimestamp,
    getCurrentRunSheet: getCurrentRunSheet,
    addAccount: addAccount,
    markAccountAsProcessed: markAccountAsProcessed,
    clearAllAccountInfo: clearAllAccountInfo,
    getUnprocessedAccountIterator: getUnprocessedAccountIterator,
    allAccountsProcessed: allAccountsProcessed,
    writeHeaderRow: writeHeaderRow,
    writeDataRow: writeDataRow
  };

})();

/**
 * Module responsible for maintaining a list of common settings. These
 * settings are read from the general spreadsheet (using the
 * spreadsheetManager) and are then retrieved by other modules during
 * processing.
 *
 * @return {object} callable functions corresponding to the available
 * actions
 */
var settingsManager = (function() {
  var settings = [];

  /**
   * Reads the settings from the general spreadsheet.
   */
  var readSettings = function() {
    var settingsRange = spreadsheetManager.readSettingRange();

    for (var i = 1; i <= settingsRange.getNumRows(); i++) {
      var key = settingsRange.getCell(i, 1).getValue();
      var type = settingsRange.getCell(i, 2).getValue();
      var value = settingsRange.getCell(i, 3).getValue();

      if (type == 'Color') {
       value = settingsRange.getCell(i, 3).getBackground();
      }

      if (!key || !value) {
        continue;
      }

      var setting = {
        key: key,
        type: type,
        value: value
      };

      settings.push(setting);
    }

    debug('Read ' + settings.length + ' settings');
  };

  /**
   * Returns the value of a particular setting.
   *
   * @param {string} key the name of the setting
   * @param {boolean} mandatory flag indicating this is a mandatory setting
   *                            (has to return a value)
   * @return {object} the value of the setting
   */
  var getSetting = function(key, mandatory) {
    for (var i = 0; i < settings.length; i++) {
      var setting = settings[i];
      if (setting.key == key && setting.value)
        return setting.value;
    }

    if (mandatory) {
      throw 'Setting \'' + key + '\' is not set!';
    }

    return null;
  };

  // Return the external interface.
  return {
    readSettings: readSettings,
    getSetting: getSetting
  };
})();

/**
 * Wrapper for Logger.log.
 *
 * @param {string} t The text to log
 */
function debug(t) {
  Logger.log(t);
}

/**
 * Validates the provided spreadsheet URL to make sure that it's set up
 * properly. Throws a descriptive error message if validation fails.
 *
 * @throws {Error} If the spreadsheet URL hasn't been set
 */
function validateConfig() {
  if (CONFIG.SPREADSHEET_URL == 'YOUR_SPREADSHEET_URL') {
    throw new Error('Please specify a valid Spreadsheet URL. You can find' +
        ' a link to a template in the associated guide for this script.');
  }
}