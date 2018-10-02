/*
 * This script runs daily to process circ transaction data to give us an idea of busy-ness trends. It:
 *
 *   ** Makes sure there's a Circ Transaction Trends file for the current year, and creates one if there isn't.
 *   ** Queries the CARL reports server for a count of transactions per hour per branch for the preceding day.
 *   ** Dumps the contents of the data array into the end of the Data tab.
 *
 * The main tab of the Circ Transaction Trends file uses dsum functions and conditional formatting to create a heat map
 * of circ transactions that can be filtered by branch and date range.
 *
 */

var address = [* Your Reports IP/port *];
var username = 'reports';
var userPwd = 'carlx';
var db = [* Your Reports DB name *];
var dbUrl = 'jdbc:oracle:thin:@//' + address + '/' + db;

function getCircTransactions() {
  
  // Look for a Drive file that has the right name; if it's not there, create it
  var now = new Date();
  var year = now.getFullYear();
  var files = DriveApp.searchFiles('title contains "' + year + ' Circ Transaction Trends"');
  var id, spreadsheet;
  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    var regExp = new RegExp(year);
    if (name.match(regExp) && name.match(/^[^(Copy)]/i)) {
      id = file.getId();
    }
  }
  if (id) {
    spreadsheet = SpreadsheetApp.openById(id);
  } else {
    spreadsheet = SpreadsheetApp.create(year + ' Circ Transaction Trends');
    spreadsheet.insertSheet('Circ Traffic Patterns', 0);
    var dataSheet = spreadsheet.insertSheet('Data', 1);
    dataSheet.getRange(1, 1, 1, 5).setValues([['Date', 'Day', 'Time', 'Branch', 'Transactions']]);
    dataSheet.deleteColumns(6, dataSheet.getLastColumn - 5);
    var emptySheet = spreadsheet.getSheets()[2];
    spreadsheet.deleteSheet(emptySheet);
  }
  
  // Pull yesterday's transaction data from CARL
  var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000)
  yesterday = getDate(yesterday);
  var today = getDate(now);
  
  /*
   * Notes about this query:
   *   ** The interior query (aliased as 't') pulls out yesterday's checkout transactions from the transaction log. If
   *      a person has more than one checkout per hour per branch, it only returns that person once.
   *   ** The CircDay element returns a number corresponding to the day of the week (1-7) that helps make the formulas in
   *      the Google Sheet work correctly.
   *   ** The CircHour element pulls the time portion of the date and then returns just the two digits corresponding to the hour.
   *   ** The <yesterday> and <today> variables in the query ensure that when this is triggered daily, it returns 
   *      yesterday's data.
   *   ** Because we use this data to get a sense of how busy we are from a staffing perspective, we're excluding non-branches
   *      from the results.
   *   ** The outer query groups the interior query results by date, hour and branch and returns a count.
   */
  var sql = 'SELECT t.circdate, t.circday, TO_NUMBER(t.circhour), t.envbranch, COUNT(t.patronid) ' +
            'FROM ' +
            '    (SELECT UNIQUE patronid, TO_CHAR(jts.todate(systemtimestamp), \'YYYY-MM-DD\') as CircDate, ' +
            '      TO_CHAR(jts.todate(systemtimestamp), \'D\') as CircDay, ' +
            '      (SELECT SUBSTR(TO_CHAR(jts.todate(systemtimestamp), \'HH24:MI\'), 0, 2) FROM DUAL) as CircHour, envbranch ' +
            '    FROM txlog_v ' +
            '    WHERE transactiontype = \'CH\' ' +
            '    AND jts.todate(systemtimestamp) > \'' + yesterday + '\' AND jts.todate(systemtimestamp) < \'' + today + '\'' +
            '    AND envbranch NOT IN (\'SE\',\'BG\',\'WV\',\'SC\',\'ON\')) t ' +
            'GROUP BY (t.circdate, t.circday, t.circhour, t.envbranch) ' +
            'ORDER BY t.circdate, t.envbranch, t.circhour';
  
  var conn = Jdbc.getConnection(dbUrl, username, userPwd);
  var stmt = conn.createStatement();
  var results = stmt.executeQuery(sql);
  var data = [];
  while (results.next()) {
    data.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5)]); 
  }
  
  // If the query returned data, append it after the last row of what's currently in the Data tab
  if (data.length > 0) {
    var sheet = spreadsheet.getSheetByName('Data');
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, data.length, 5).setValues(data);
  }
}


/*
 * @params {date} date Javascript date passed in from main function for conversion
 * @return {string} converted Julian Time Stamp date 
 */
function getDate(date) {
  // Get the month (0-11) from the Javascript date, and match to the 3-letter month abbreviation
  var textMonth;
  var month = date.getMonth();
  if (month == 0) {
    textMonth = 'JAN';
  } else if (month == 1) {
    textMonth = 'FEB';
  } else if (month == 2) {
    textMonth = 'MAR';
  } else if (month == 3) {
    textMonth = 'APR';
  } else if (month == 4) {
    textMonth = 'MAY';
  } else if (month == 5) {
    textMonth = 'JUN';
  } else if (month == 6) {
    textMonth = 'JUL';
  } else if (month == 7) {
    textMonth = 'AUG';
  } else if (month == 8) {
    textMonth = 'SEP';
  } else if (month == 9) {
    textMonth = 'OCT';
  } else if (month == 10) {
    textMonth = 'NOV';
  } else if (month == 11) {
    textMonth = 'DEC';
  }
  // Get the year and date
  var year = date.getFullYear();
  var day = date.getDate();
  
  // If the day is less than 10, put a 0 at the beginning of the date; otherwise, return as-is
  if (day < 10) {
    return '0' + day.toString() + '-' + textMonth + '-' + year;
  } else {
    return day.toString() + '-' + textMonth + '-' + year;
  }
}
  
