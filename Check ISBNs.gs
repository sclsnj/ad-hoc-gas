/*
 * This script can be deployed as a Google Sheets add on to check a set of existing ISBNs against the database
 * and return matching BIDs.  
 */

/*
 * This sets the trigger so that the user sees the appropriate menu item if they have the add on installed.
 * @params {object} e the spreadsheet object being opened
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Check ISBNs', 'readyCheck')
      .addToUi();
}

/*
 * This sets the trigger so that the appropriate menu item is added when the user first installs the add on.
 * @params {object} e the spreadsheet object being opened
 */
function onInstall(e) {
  onOpen(e);
}


/*
 * This function runs whenever the user selects "Check ISBNs" from the add on menu, to prompt to ensure that
 * ISBN data for checking is in the right place. If the user says no, the checkISBNs function doesn't run.
 */

function readyCheck() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Ready to check?', 'Your list of ISBNs to check should be in the first column, with ' + 
            'no header and no other data. \n\n' +
            'Are your ISBNs ready to go?', ui.ButtonSet.YES_NO);
  if (response == 'YES') {
    checkISBNs();
  }
}

var address = [* Your Reports IP/port *];
var username = 'reports';
var userPwd = 'carlx';
var db = [* Your Reports DB name *];
var dbUrl = 'jdbc:oracle:thin:@//' + address + '/' + db;


function checkISBNs() {
  var conn = Jdbc.getConnection(dbUrl, username, userPwd);
  var stmt = conn.createStatement();
  
  // Get the current spreadsheet, and pull the existing ISBNs into an array we can loop through.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getDataRange();
  var isbns = range.getValues();
  var exists = [];
  
  // For each ISBN, run the query to find a matching BID, and then capture the ISBN and BID information.
  for (var i = 0; i < isbns.length; i++) {
    var isbn = isbns[i][0].toString();
    isbn = isbn.replace(/[^0-9X]/i, '');

    /*
     * Notes about this query:
     *   ** This returns any BIDs that have something that looks like the existing ISBN in the 020, 024 or 028 tags.
     */

    var sql = 'SELECT b.bid ' +
              'FROM bbibmap_v b, btags_v t, bbibcontents_v bc ' +
              'WHERE bc.tagid = t.tagid AND bc.bid = b.bid ' +
              'AND (bc.tagnumber = \'028\' OR bc.tagnumber = \'020\' OR bc.tagnumber = \'024\') ' +
              'AND t.worddata LIKE \'' + isbn + '%\'';

    var results = stmt.executeQuery(sql);
    var data = [];
    if (results) {
      while (results.next()) {
        data.push([results.getString(1), results.getString(2), results.getString(3)]);
      }
    }
    
    // If no BIDs come back from the query, put just the ISBN back in the array. Otherwise, put the ISBN and the first
    // matching BID in the array.
    
    if (data.length == 0 || !data) {
      exists.push([isbn, '']);
    } else {
      exists.push([isbn, 'BID ' + data[0][0]]);
    }
  }
  
  // Replace the ISBNs that are in the sheet with the new set of ISBNs and matching BIDs.
  range = sheet.getRange(1, 1, exists.length, 2);
  range.setValues(exists);
}
