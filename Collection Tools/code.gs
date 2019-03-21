/*

This project is for a standalone add-on for Google Sheets that allows Collection Development 
to automate running and formatting collection reports.

The add-on has been published in the Chrome Web Store and made available only to SCLSNJ
domain users.

This Code.gs file includes the Javascript tied to CARLcollchecksidebar.html, as well as utility functions
called by the CARLcode.gs.

*/

// Immediately initiates the menu when the add-on is installed
function onInstall(e) {
  onOpen(e);
}

// Initiates the add-on menu to include the available functions whenever a Sheet is opened
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('CARL Collection Check', 'showCARLSidebar')
    .addToUi();
}


// Opens the 'Format Collection Check' sidebar as html
function showCARLSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('CARLcollchecksidebar')
      .setTitle('CARL Collection Check')
  SpreadsheetApp.getUi().showSidebar(ui);
}

/*
 * All the rest of this code gets invoked after the user chooses variables from the sidebar and submits a request.
 */

// Called from CARLcollchecksidebar; gets the stored user preferences for formatting collection checks, if they exist
function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  var formatPrefs = {
    useOldandout: userProperties.getProperty('useOldandout'),
    pubdateAge: userProperties.getProperty('pubdateAge'),
    useOldorhold: userProperties.getProperty('useOldorhold'),
    useNocirc: userProperties.getProperty('useNocirc'),
    ladAge: userProperties.getProperty('ladAge'),
    useOldwithrecent: userProperties.getProperty('useOldwithrecent'),
    createdAge: userProperties.getProperty('createdAge'),
    useSupercirc: userProperties.getProperty('useSupercirc'),
    bigCirc: userProperties.getProperty('bigCirc'),
    useConstantcirc: userProperties.getProperty('useConstantcirc'),
    circRate: userProperties.getProperty('circRate')
  };
  return formatPrefs;
}


// formatReport Utility: formats report columns, rows, etc., sets conditional formatting for flags
// NOTE: This is a generic function that's used in other SCLSNJ add ons, which is why there are columns, flags and types listed here
//       that don't correspond to the data coming back from the collection check. -LH
function formatReport(sheet) {
  var numRows = sheet.getLastRow();
  var numCols = sheet.getLastColumn();
  var values = sheet.getDataRange().getValues();
  var dataType, width, heat = '';
  var hide, wrap = false;
  var colFormats = [];
  
  // Create an array containing formatting criteria for each data type, column width, heat flag, column hide, text wrap 
  for (var col = 0; col < numCols; col++) {
    var testHeader = values[0][col].toString();
    if (testHeader.match('Call Range')) {
      width = '162';
      wrap = true;
    }
    if (testHeader.match(/call/i) || testHeader.match(/barcode/i)) {
      dataType = '@STRING@';
    }
    else if (testHeader.match(/subject/i) || testHeader.match(/title/i) || testHeader.match(/author/i)) {
      dataType = '@STRING@';
    }
    else if (testHeader.match(/pub date/i) || testHeader.match(/stock/i)) {
      dataType = '@STRING@';
    }
    else if (testHeader.match(/% /)) {
      dataType = '0.00%';
    }
    else if (testHeader.match(/total/i) || testHeader.match(/chkout/i) || testHeader.match(/circ/i) || testHeader.match(/use/i)) {
      dataType = '0';
    }
    else if (testHeader.match(/turnover/i)) {
      dataType = '0.00';
      width = '90';
      heat = 'turnover';
    }
    else if (testHeader.match('Stock Level')) {
      width = '90';
      heat = 'overUnder';
    }
    else if (testHeader.match(/created/i) || testHeader.match(/date/i)) {
      dataType = 'mm/dd/yyyy';
    }
    else if (testHeader.match('Subject')) {
      width = '192';
      wrap = true;
    }
    else if (testHeader.match(/call #/i)) {
      width = '100';
    }
    else if (testHeader.match(/location/i)) {
      hide = true;
    }
    else if (testHeader.match(/title/i)) {
      width = '300';
    }
    else if (testHeader.match(/author/i)) {
      width = '200';
    }
    else if (testHeader.match(/barcode/i) || testHeader.match(/item/i)) {
      width = '120';
    }
    else if ((testHeader.match(/chkout/i)) || (testHeader.match(/circ/i))) {
      width = '40';
    }
    else if (testHeader.match(/flag/i)) {
      width = '235';
      heat = 'flag';
    }
    else {
      width = '70';
    }
    colFormats.push( [ col, width, heat, hide, wrap ] );
    dataType = '';
    width = '70';
    heat = '';
    wrap = false;
    hide = false;
  }

  // Parse through the formating instructions for each column and set the sheet accordingly
  for (var col = 0; col < numCols; col++) {
    var range = sheet.getRange(1, col + 1, numRows);    
    sheet.setColumnWidth(col + 1, colFormats[col][1]);               // sets column width
    if (colFormats[col][3]) {                                        // hides column, if specified
      sheet.hideColumns(col + 1);
    }
    if (colFormats[col][4]) {                                        // wraps text, if specified
      range.setWrap(true);
    }    
    
    // These first two conditional formatting pieces are left over from another utility; the one that's applicable here is 
    // the third one: else if (colFormats[col][2] == 'flag')
    if (colFormats[col][2] == 'overUnder') {                         // sets conditional formatting based on column data
      var overUnder = sheet.getRange(1, col + 1, numRows).getValues(); 
      for (i = 1; i < overUnder.length; i++) {
        if (overUnder[i][0].toString() == 'understocked') {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackgroundRGB(183,225,205);
        }
        else if (overUnder[i][0].toString() == 'overstocked') {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackgroundRGB(244,199,195);
        }
      }
    }
    else if (colFormats[col][2] == 'turnover') { 
      var turnover = sheet.getRange(1, col + 1, numRows).getValues();
      for (var i = 1; i < turnover.length; i++) {
        if (turnover[i][0] >= 1) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackgroundRGB(183,225,205);
        }
        else if (turnover[i][0] < 0.6) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackgroundRGB(244,199,195);
        }
      }
    }
    // NOTE: For what it's worth, there's probably a better way to do this using some actual conditional formatting
    //       methods in GAS that would take less overhead than hard setting the background color of individual cells
    //       based on their actual values.
    else if (colFormats[col][2] == 'flag') { 
      var flags = sheet.getRange(1, col + 1, numRows).getValues();
      for (var i = 1; i < flags.length; i++) {
        var flag = flags[i][0].toString();
        if (flag.match('Checked out, older pub - replace?')) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackground('#fff2cc');
        }
        else if (flag.match('Checked out or on hold')) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackground('#f4cccc');
        }
        else if (flag.match('Zero circs')) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackground('#cfe2f3');
        }
        else if (flag.match('No recent activity - deselect?')) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackground('#d9ead3');
        }
        else if (flag.match('years old - deselect or replace?')) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackground('#fce5cd');
        }
        else if (flag.match('circs - deselect or replace?')) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackground('#ead1dc');
        }
        else if (flag.match('circs per year')) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackground('#d9d2e9');
        }
        else if (flag.match(/missing/i) || flag.match(/lost/i) || flag.match(/search/i) || flag.match(/claims/i) || flag.match(/billed/i) || flag.match(/found/i)) {
          sheet.getRange(i + 1, col + 1, 1, 1).setBackground('#d9d9d9');
        }
      }
    }
  }
  
  // Once all the columns have been dealt with, set other range- and row-related formatting options.
  for (var row = 1; row <= numRows; row++) {          // sets consistent row height
    sheet.setRowHeight(row, 28);
  }
  range = sheet.getDataRange();                       // sets alignment, font and size
  range
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setFontSize('11')
    .setFontFamily('Arial'); 
  range = sheet.getRange(1, 1, 1, col);               // formats first row: string, v-align, h-align, bold
  range
    .setNumberFormat('@STRING@')
    .setVerticalAlignment('bottom')
    .setHorizontalAlignment('left')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);                             // freezes first row
}
