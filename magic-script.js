/**
 * A script to automate requesting data from Google Analytics.
 *
 * @author nickski15@gmail.com (Nick Mihailovski)
 */


/**
 * The name of the configration sheet.
 * And various parameters.
 */
var GA_CONFIG = 'gaconfig';
var NAME = 'query';
var VALUE = 'value';
var TYPE = 'type';
var SHEET_NAME = 'sheet-name';

var CORE_OPT_PARAM_NAMES = [
  'dimensions',
  'sort',
  'filters',
  'segment',
  'start-index',
  'max-results'];

var MCF_OPT_PARAM_NAMES = [
  'dimensions',
  'sort',
  'filters',
  'start-index',
  'max-results'];


/**
 * The types of reports.
 */
var CORE_TYPE = 'core';
var MCF_TYPE = 'mcf';


/**
 * Create a Menu when the script loads. Adds a new gaconfig sheet if
 * one doesn't exist.
 */
function onOpen() {

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Add a menu.
  activeSpreadsheet.addMenu(
      'Google Analytics', [{
        name: 'Find Profile / ids', functionName: 'findIds_'}, {
        name: 'Create Core Report', functionName: 'createCoreReport'}, {
        name: 'Create MCF Report', functionName: 'createMcfReport'}, {
        name: 'Get Data', functionName: 'getData'
      }]);

  // Add a sheet called gaconfig if it doesn't already exist.
  getOrCreateGaSheet_();
}


/**
 * Opens a UIService dialogue for users to select their Profile / ids
 * query parameter. This will traverse the Management API hierarchy
 * to populate various list boxes.
 */
function findIds_() {
  var app = UiApp.createApplication()
      .setTitle('Find Your Profile ID / ids Parameter')
      .setWidth(700).setHeight(175);

  var accountList = app.createListBox()
      .setId('accountList').setName('accountList')
      .setVisibleItemCount(1).setStyleAttribute('width', '100%');
  var webpropertyList = app.createListBox()
      .setId('webpropertyList').setName('webpropertyList')
      .setVisibleItemCount(1).setStyleAttribute('width', '100%');
  var profileList = app.createListBox()
      .setId('profileList').setName('profileList')
      .setVisibleItemCount(1).setStyleAttribute('width', '100%');

  // Grid for dropdowns.
  var grid = app.createGrid(3, 2).setStyleAttribute(0, 0, 'width', '130px');
  grid.setWidget(0, 0, app.createLabel('Select Account'));
  grid.setWidget(0, 1, accountList);
  grid.setWidget(1, 0, app.createLabel('Select Web Property'));
  grid.setWidget(1, 1, webpropertyList);
  grid.setWidget(2, 0, app.createLabel('Select Profile'));
  grid.setWidget(2, 1, profileList);

  grid.setStyleAttribute('border-bottom', '1px solid #999')
      .setStyleAttribute('padding-bottom', '12px')
      .setStyleAttribute('margin-bottom', '12px');

  // Grid for id results.

  //var profileLink = app.createLabel('add ids to sheet').addClickHandler();

  var grid2 = app.createGrid(2, 3).setStyleAttribute(0, 0, 'width', '130px');
  grid2.setWidget(0, 0, app.createLabel('Profile Id'));
  grid2.setWidget(0, 1, app.createLabel().setId('profileId'));
  grid2.setWidget(1, 0, app.createLabel('ids'));
  grid2.setWidget(1, 1, app.createLabel().setId('ids'));
  //grid2.setWidget(1, 2, profileLink);

  var handler = app.createServerHandler('queryWebProperties_');
  accountList.addChangeHandler(handler);

  handler = app.createServerHandler('queryProfiles_');
  webpropertyList.addChangeHandler(handler);

  handler = app.createServerHandler('displayProfile_');
  profileList.addChangeHandler(handler);

  app.add(grid).add(grid2);

  // Traverse the management hiearchy by default.
  queryAccounts_();

  // Show display.
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  doc.show(app);
}


/**
 * Handler to query all the accounts for the user then update the
 * UI dialogue box.
 * @return {Object} The instance of the active application.
 */
function queryAccounts_() {
  var app = UiApp.getActiveApplication();
  var accountList = app.getElementById('accountList');

  // Query Accounts API
  var accounts = Analytics.Management.Accounts.list();

  if (accounts.getItems()) {
    for (var i = 0, item; item = accounts.getItems()[i]; ++i) {
      var name = item.getName();
      var id = item.getId();
      accountList.addItem(name, id);
    }

    var firstAccountId = accounts.getItems()[0].getId();
    UserProperties.setProperty('accountId', firstAccountId);
    queryWebProperties_();

  } else {
    accountList.addItem('No accounts for user.');
    app.getElementById('webpropertyList').clear();
    app.getElementById('profileList').clear();
    displayProfile_('none');
  }

  return app;
}


/**
 * Handler to query all the webproperties.
 * @param {object} eventInfo Used retrieve the user's accountId.
 * @return {object} A reference to the active application.
 */
function queryWebProperties_(eventInfo) {
  var app = UiApp.getActiveApplication();
  var webpropertyList = app.getElementById('webpropertyList').clear();

  // From saved id.
  var accountId = UserProperties.getProperty('accountId');

  // From server handler.
  if (eventInfo && eventInfo.parameter && eventInfo.parameter.accountList) {
    accountId = eventInfo.parameter.accountList;
    UserProperties.setProperty('accountId', accountId);
  }

  var webproperties = Analytics.Management.Webproperties.list(accountId);
  if (webproperties.getItems()) {
    for (var i = 0, webproperty; webproperty = webproperties.getItems()[i];
        ++i) {

      var name = webproperty.getName();
      var id = webproperty.getId();
      webpropertyList.addItem(name, id);
    }

    var firstWebpropertyId = webproperties.getItems()[0].getId();
    UserProperties.setProperty('webpropertyId', firstWebpropertyId);
    queryProfiles_();

  } else {
    webpropertyList.addItem('No webproperties for user');
    app.getElementById('profileList').clear();
    displayProfile_('none');
  }

  return app;
}


/**
 * Handler to query all the profiles.
 * @param {object} eventInfo Used retrieve the user's webpropertyId.
 * @return {object} A reference to the active application.
 */
function queryProfiles_(eventInfo) {
  var app = UiApp.getActiveApplication();
  var profileList = app.getElementById('profileList').clear();

  var accountId = UserProperties.getProperty('accountId');
  var webpropertyId = UserProperties.getProperty('webpropertyId');
  if (eventInfo && eventInfo.parameter && eventInfo.parameter.webpropertyList) {
    webpropertyId = eventInfo.parameter.webpropertyList;
  }

  var profiles = Analytics.Management.Profiles.list(accountId, webpropertyId);
  if (profiles.getItems()) {
    for (var i = 0, profile; profile = profiles.getItems()[i]; ++i) {
      profileList.addItem(profile.getName(), profile.getId());
    }

    var firstProfileId = profiles.getItems()[0].getId();
    UserProperties.setProperty('profileId', firstProfileId);
    displayProfile_();

  } else {
    profileList.addItem('No profiles found.');
    displayProfile_('none');
  }
  return app;
}


/**
 * Displays the user's profile on the spreadsheet.
 * @param {Objecty} eventInfo Used to retrieve the profile ID.
 * @return {object} A reference to the active application.
 */
function displayProfile_(eventInfo) {
  var app = UiApp.getActiveApplication();

  var profileId = '';
  var tableId = '';

  if (eventInfo != 'none') {
    var profileId = UserProperties.getProperty('profileId');
    if (eventInfo && eventInfo.parameter && eventInfo.parameter.profileList) {
      profileId = eventInfo.parameter.profileList;
    }
    var tableId = 'ga:' + profileId;
  }

  app.getElementById('profileId').setText(profileId);
  app.getElementById('ids').setText(tableId);
  return app;
}


/**
 * Returns the GA_CONFIG sheet. If it doesn't exist it is created.
 * @return {Sheet} The GA_CONFIG sheet.
 */
function getOrCreateGaSheet_() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var gaConfig = activeSpreadsheet.getSheetByName(GA_CONFIG);
  if (!gaConfig) {
    gaConfig = activeSpreadsheet.insertSheet(GA_CONFIG, 0);
  }
  return gaConfig;
}


/**
 * Entry point for a user to create a new Core Report configuration.
 */
function createCoreReport() {
  createGaReport_(CORE_TYPE);
}


/**
 * Entry point for a user to create a new MCF Report Configuration.
 */
function createMcfReport() {
  createGaReport_(MCF_TYPE);
}


/**
 * Adds a GA Report configuration to the spreadsheet.
 * @param {String} reportType The type of report configuration to add.
 *    Should be either CORE_TYPE or MCF_TYPE.
 */
function createGaReport_(reportType) {
  var sheet = getOrCreateGaSheet_();
  var headerNumber = getLastNumber_(sheet);
  var config = [
    [NAME + headerNumber, VALUE + headerNumber],
    [TYPE, reportType],
    ['ids', ''],
    ['start-date', ''],
    ['end-date', ''],
    ['last-n-days', ''],
    ['metrics', '']];

  // Add API specific fields. Default to Core.
  var paramNames = CORE_OPT_PARAM_NAMES;
  if (reportType == MCF_TYPE) {
    paramNames = MCF_OPT_PARAM_NAMES;
  }
  for (var i = 0, paramName; paramName = paramNames[i]; ++i) {
    config.push([paramName, '']);
  }

  config.push([SHEET_NAME, '']);
  sheet.getRange(1, sheet.getLastColumn() + 1, config.length, 2)
       .setValues(config);
}


/**
 * Returns 1 greater than the largest trailing number in the header row.
 * @param {Object} sheet The sheet in which to find the last number.
 * @return {Number} The next largest trailing number.
 */
function getLastNumber_(sheet) {
  var maxNumber = 0;

  var lastColIndex = sheet.getLastColumn();

  if (lastColIndex > 0) {
    var range = sheet.getRange(1, 1, 1, lastColIndex);

    for (var colIndex = 1; colIndex < sheet.getLastColumn(); ++colIndex) {
      var value = range.getCell(1, colIndex).getValue();
      var headerNumber = getTrailingNumber_(value);
      if (headerNumber) {
        var number = parseInt(headerNumber, 10);
        maxNumber = number > maxNumber ? number : maxNumber;
      }
    }
  }
  return maxNumber + 1;
}


/**
 * Main function to get data from the Analytics API. This reads the
 * configuration file, executes the queries, and displays the results.
 * @param {object} e Some undocumented parameter that is passed only
 *     when the script is run from a trigger. Used to update
 *     user on the status of the report.
 */
function getData(e) {
  setupLog_();
  var now = new Date();
  log_('Running on: ' + now);

  var sheet = getOrCreateGaSheet_();
  var configs = getConfigs_(sheet);

  if (!configs.length) {
    log_('No report configurations found');

  } else {
    log_('Found ' + configs.length + ' report configurations.');

    for (var i = 0, config; config = configs[i]; ++i) {
      var configName = config[NAME];

      try {
        log_('Executing query: ' + configName);
        var results = getResults_(config);

        log_('Success. Writing results.');
        displayResults_(results, config);

      } catch (error) {
        log_('Error executing ' + configName + ': ' + error.message);
      }
    }
  }
  log_('Script done');

  // Update the user about the status of the queries.
  if (e === undefined) {
    displayLog_();
  }
}


/**
 * Returns an array of config objects. This reads the gaconfig sheet
 * and tries to extract adjacent column names that end with the same
 * number. For example Names1 : Values1. Then both columns are used
 * to define key-value pairs for the coniguration object. The first
 * column defines the keys, and the adjacent column values define
 * each keys values.
 * @param {Sheet} sheet The GA_Config sheet from which to read configurations.
 * @return {Array} An array of API query configuration object.
 */
function getConfigs_(sheet) {

  var configs = [];
  // There must be at least 2 columns.
  if (sheet.getLastColumn() < 2) {
    return configs;
  }

  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());

  // Test the name of each column to see if it has an adjacent column that ends
  // in the same number. ie xxxx555 : yyyy555.
  // Since we check 2 columns at a time, we don't need to check the last column,
  // as there is no second column to also check.
  for (var colIndex = 1; colIndex <= headerRange.getNumColumns() - 1;
      ++colIndex) {

    var firstColValue = headerRange.getCell(1, colIndex).getValue();
    var firstColNum = getTrailingNumber_(firstColValue);

    var secondColValue = headerRange.getCell(1, colIndex + 1).getValue();
    var secondColNum = getTrailingNumber_(secondColValue);

    if (firstColNum && secondColNum && firstColNum == secondColNum) {
      configs.push(getConfigsStartingAtCol_(colIndex));
    }
  }

  return configs;
}


/**
 * Returns the trailing number on a string. For example the
 * input: xxxx555 will return 555. Inputs with no trailing numbers
 * return undefined. Trailing whitespace is not ignored.
 * @param {string} input The input to parse.
 * @return {number} The trailing number on the input as a string.
 *     undefined if no number was found.
 */
function getTrailingNumber_(input) {
  // Match at one or more digits at the end of the string.
  var pattern = /(\d+)$/;
  var result = pattern.exec(input);
  if (result) {
    // Return the matched number.
    return result[0];
  }

  return undefined;
}


/**
 * Returns the values from 2 columns from the GA_CONFIG sheet starting at
 * colIndex, as key-value pairs. Key-values are only returned if they do
 * not contain the empty string or have a boolean value of false.
 * If the key is start-date or end-date and the value is an instance of
 * the date object, the value will be converted to a string in yyyy-MM-dd.
 * If the key is start-index or max-results and the type of the value is
 * number, the value will be parsed into a string.
 * @param {number} colIndex The column index to return values from.
 * @return {object} The values starting in colIndex and the following column
       as key-value pairs.
 */
function getConfigsStartingAtCol_(colIndex) {
  var config = {};

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GA_CONFIG);
  var range = sheet.getRange(1, colIndex, sheet.getLastRow(), 2);

  // The first cell of the first column becomes the name of the query.
  config[NAME] = range.getCell(1, 1).getValue();

  for (var rowIndex = 2; rowIndex <= range.getLastRow(); ++rowIndex) {
    var key = range.getCell(rowIndex, 1).getValue();
    var value = range.getCell(rowIndex, 2).getValue();
    if (value) {
      if ((key == 'start-date' || key == 'end-date') && value instanceof Date) {
        // Utilities.formatDate is too complicated since it requires a time zone
        // which can be configured by account or per sheet.
        value = formatGaDate_(value);

      } else if ((key == 'start-index' || key == 'max-results') &&
          typeof value == 'number') {

        value = value.toString();
      }
      config[key] = value;
    }
  }

  return config;
}


/**
 * Returns the dateInput object in yyyy-MM-dd.
 * @param {Date} inputDate The object to convert.
 * @return {string} The date object as yyyy-MM-dd.
 */
function formatGaDate_(inputDate) {
  var output = [];
  var year = inputDate.getFullYear();

  var month = inputDate.getMonth() + 1;
  if (month < 10) {
    month = '0' + month;
  }

  var day = inputDate.getDate();
  if (day < 10) {
    day = '0' + day;
  }
  return [year, month, day].join('-');
}


/**
 * Executes a Google Analytics API query and returns the results.
 * @param {object} config A configuration object with key-value
 *     parameters representing various API query parameters.
 * @return {object} A JSON object with all the results from the API.
 */
function getResults_(config) {

  // Translate last n days into the actual start and end dates.
  // This value overrides any existing start or end dates.
  if (config['last-n-days']) {
    var lastNdays = parseInt(config['last-n-days'], 10);
    config['start-date'] = getLastNdays_(lastNdays);
    config['end-date'] = getLastNdays_(0);
  }

  var type = config[TYPE] || CORE_TYPE;  // If no type, default to core type.
  if (type == CORE_TYPE) {
    var optParameters = getOptParamObject_(CORE_OPT_PARAM_NAMES, config);
    var apiFunction = Analytics.Data.Ga.get;

  } else if (type == MCF_TYPE) {
    var optParameters = getOptParamObject_(MCF_OPT_PARAM_NAMES, config);
    var apiFunction = Analytics.Data.Mcf.get;
  }

  // Execute query and return the results.
  // If any errors occur, they will be thrown and caught in a higher
  // level of code.
  var results = apiFunction(
      config['ids'],
      config['start-date'],
      config['end-date'],
      config['metrics'],
      optParameters);

  return results;
}


/**
 * Returns a date formatted as YYYY-MM-DD from the number of
 * days ago starting from today.
 * @param {number} nDays The number of days to request dats from today.
 * @return {String} The YYY_MM_DD formatted date of the previous days.
 */
function getLastNdays_(nDays) {
  var today = new Date();
  var before = new Date();
  before.setDate(today.getDate() - nDays);
  return formatGaDate_(before);
}


/**
 * Returns all the valid optional parameters for a Reporting API Query.
 * This ensures non-valid parameters are not added to the query.
 * Parameters with empty values will not be added.
 * @param {array.<String>} optParamNames An array of all the valid param names.
 * @param {Object} config The configuration object.
 * @return {object} An object with all the keys as param names and values as
 *     param values.
 */
function getOptParamObject_(optParamNames, config) {
  var optParameters = {};
  for (var i = 0, paramName; paramName = optParamNames[i]; ++i) {
    if (paramName in config && config[paramName]) {
      optParameters[paramName] = config[paramName];
    }
  }
  return optParameters;
}


/**
 * Displays all the API results of a sucessful query.
 * @param {object} results The data returned from the API.
 * @param {object} config An object that contains key value configuration
 *     parameters.
 */
function displayResults_(results, config) {
  // Use an object to force passing by reference.
  var row = {};
  row.count = 1;

  var activeSheet = getOutputSheet_(config[SHEET_NAME]);
  activeSheet.clear().setColumnWidth(1, 200);

  outputResultInfo_(results, activeSheet, row, config[NAME]);
  outputTotalsForAllResults_(results, activeSheet, row);
  outputHeaders_(results, activeSheet, row);

  // Only output rows if they exist.
  if (results.getRows()) {
    var type = config[TYPE] || CORE_TYPE;  // If no type, default to core type.
    if (type == CORE_TYPE) {
      outputCoreRows_(results, activeSheet, row);
    } else if (type == MCF_TYPE) {
      outputMcfRows_(results, activeSheet, row);
    }
  }
}


/**
 * Returns the sheet which should be used to output the results.
 * If no sheetName is used, a new sheet will be inserted.
 * If a sheetName is used, but doesn't yet exist, it will be inserted.
 * If a sheetName is used, and already exists, it will be returned.
 * @param {string=} opt_sheetName The name of the sheet to return.
 * @return {object} A reference to the active spreadsheet.
 */
function getOutputSheet_(opt_sheetName) {
  var sheetName = opt_sheetName || false;
  var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet;
  if (sheetName) {
    sheetName += ''; // Make sure sheet is a string.
    activeSheet = activeSpreadSheet.getSheetByName(sheetName);
    if (!activeSheet) {
      activeSheet = activeSpreadSheet.insertSheet(sheetName);
    }
  } else {
    activeSheet = activeSpreadSheet.insertSheet();
  }
  return activeSheet;
}


/**
 * Outputs whether the result contains sampled data.
 * @param {object} results The object returned from the API.
 * @param {object} activeSheet The active sheet to display the results.
 * @param {object} row An object that stores on which row to start outputing.
 * @param {String} queryName The name of the query.
 */
function outputResultInfo_(results, activeSheet, row, queryName) {
  var now = Utilities.formatDate(new Date(), Session.getTimeZone(),
      'yyyy-MM-dd HH:mm:ss');

  activeSheet.getRange(row.count, 1, 6, 2).setValues([
    ['Results for query', queryName],
    ['Date executed', now],
    ['Profile Name', results.getProfileInfo().getProfileName()],
    ['Total Results Found', results.getTotalResults()],
    ['Total Results Returned', results.getRows().length],
    ['Contains Sampled Data', results.getContainsSampledData()]
  ]);

  // Merge 4 cells to make it look nice.
  // For date / time.
  activeSheet.getRange(row.count + 1, 2, 1, 2).mergeAcross();
  // For profile name.
  activeSheet.getRange(row.count + 2, 2, 1, 4).mergeAcross();
  row.count += 7;
}


/**
 * Outpus the totals for all results. Goes through each header in one pass.
 * The number of dimensions are counted so that the range can be aligned with
 * the headers for all the rows. The metric names are put into an array in the
 * order they appear in the results. The totals for each metrics are also
 * looked up and put into an array in the same order as each name. Finally,
 * the name and totals are outputted into the sheet.
 * @param {object} results The object returned from the API.
 * @param {object} activeSheet The active sheet to display the results.
 * @param {object} row An object that stores on which row to start outputing.
 */
function outputTotalsForAllResults_(results, activeSheet, row) {
  activeSheet.getRange(row.count, 1).setValue('Totals For All Results');
  ++row.count;

  var numDimensions = 0;
  var metricNames = [];
  var metricTotals = [];
  for (var i = 0, header; header = results.getColumnHeaders()[i]; ++i) {
    if (header.getColumnType() == 'DIMENSION') {
      ++numDimensions;
    } else {
      var metricName = header.getName();
      metricNames.push(metricName);
      var totalForMetric = results.getTotalsForAllResults()[metricName];
      metricTotals.push(totalForMetric);
    }
  }

  // Get a range that skips over the dimensions.
  var range = activeSheet.getRange(row.count, numDimensions + 1, 2,
                                   results.getColumnHeaders().length -
                                       numDimensions);
  range.setValues([metricNames, metricTotals]);
  row.count += 3;
}


/**
 * Outputs the header values in the activeSheet.
 * @param {object} results The object returned from the API.
 * @param {object} activeSheet The active sheet to display the results.
 * @param {object} row An object that stores on which row to start outputing.
 */
function outputHeaders_(results, activeSheet, row) {
  var headerRange = activeSheet.getRange(row.count, 1, 1,
      results.getColumnHeaders().length);
  var headerNames = [];
  for (var i = 0; i < results.getColumnHeaders().length; ++i) {
    headerNames.push(results.getColumnHeaders()[i].getName());
  }
  headerRange.setValues([headerNames]);

  row.count += 1;
}


/**
 * Outputs the rows of data in the activeSheet. This also updates the first
 * occurance of the column that has date data to be in yyyy=MM-dd format.
 * @param {object} results The object returned from the API.
 * @param {object} activeSheet The active sheet to display the results.
 * @param {object} row An object that stores on which row to start outputing.
 */
function outputCoreRows_(results, activeSheet, row) {
  var rows = results.getRows();

  // Update the ga:date columns to be in yyyy-MM-dd format.
  var dateCols = getDateColumns_(results.getColumnHeaders());
  if (dateCols.length) {

    for (var rowIndex = 0; rowIndex < rows.length; ++rowIndex) {
      for (var i = 0; i < dateCols.length; ++i) {
        var dateColIndex = dateCols[i];
        rows[rowIndex][dateColIndex] = convertToIsoDate_(
            rows[rowIndex][dateColIndex]);
      }
    }
  }

  activeSheet.getRange(row.count, 1,
                       results.getRows().length,
                       results.getColumnHeaders().length).setValues(rows);
}


/**
 * Outputs the rows of data in the active sheet. Each primitive dimension gets
 * it's own cell. Each primitiveValue gets it's own cell. The nodes of each
 * conversionPathValue get individual cells. Because the size of the resulting
 * table may vary, a counter keeps track of the final table.
 * @param {object} results The object returned from the API.
 * @param {object} activeSheet The active sheet to display the results.
 * @param {object} row An object that stores on which row to start outputing.
 */
function outputMcfRows_(results, activeSheet, row) {
  var maxCols = 1;
  var outputTable = [];
  for (var rowIndex = 0, resultRow; resultRow = results.getRows()[rowIndex];
      ++rowIndex) {

    var outputRow = [];
    for (var colIndex = 0, cell; cell = resultRow[colIndex]; ++colIndex) {
      var jsonCell = Utilities.jsonParse(cell);

      if (jsonCell['primitiveValue']) {
        outputRow.push(jsonCell['primitiveValue']);

      } else if (jsonCell['conversionPathValue']) {
        // All nodes become cells.
        var nodeCell = [];
        var nodes = jsonCell['conversionPathValue'];
        for (var nodeIndex = 0, node; node = nodes[nodeIndex]; ++nodeIndex) {
          nodeCell.push(node['nodeValue']);
        }
        outputRow.push(nodeCell.join(' > '));  //TODO: should we auto-expand?
      }
    }
    outputTable.push(outputRow);
    if (outputRow.length > maxCols) {
      maxCols = outputRow.length;
    }
  }

  var range = activeSheet.getRange(row.count, 1,
                                   outputTable.length,
                                   maxCols);
  range.setValues(outputTable);
}


/**
 * Returns an array with all the indicies of columns that contain the ga:date
 * dimension. The indicies start from 0.
 * @param {object} headers The header object returned from the API.
 * @return {Array.<number>} The 0-based indicies of the columns that contain.
 */
function getDateColumns_(headers) {
  var indicies = [];
  for (var colIndex = 0; colIndex < headers.length; ++colIndex) {
    var name = headers[colIndex].getName();
    if (name == 'ga:date') {
      indicies.push(colIndex);
    }
  }
  return indicies;
}


/**
 * Converts a string in the format yyyyMMdd to the format yyyy-MM-dd.
 * @param {String} gaDate The string to convert.
 * @return {String} The formatted string.
 */
function convertToIsoDate_(gaDate) {
  var year = gaDate.substring(0, 4);
  var month = gaDate.substring(4, 6);
  var day = gaDate.substring(6, 8);
  return [year, month, day].join('-');
}


/**
 * The output text that should be displayed in the log.
 * @private.
 */
var logArray_;


/**
 * Clears the in app log.
 * @private.
 */
function setupLog_() {
  logArray_ = [];
}


/**
 * Appends a string as a new line to the log.
 * @param {String} value The value to add to the log.
 */
function log_(value) {
  logArray_.push(value);

  var app = UiApp.getActiveApplication();
  var foo = app.getElementById('log');
  foo.setText(getLog_());
}


/**
 * Returns the log as a string.
 * @return {string} The log.
 */
function getLog_() {
  return logArray_.join('\n');
}


/**
 * Displays the log in memory to the user.
 */
function displayLog_() {
  var uiLog = UiApp.createApplication().setTitle('Report Status')
                   .setWidth(400).setHeight(500);
  var panel = uiLog.createVerticalPanel();
  uiLog.add(panel);

  var txtOutput = uiLog.createTextArea().setId('log').setWidth('400')
                       .setHeight('500').setValue(getLog_());
  panel.add(txtOutput);

  SpreadsheetApp.getActiveSpreadsheet().show(uiLog);
}
