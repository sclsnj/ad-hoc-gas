/*
 * This script runs weekly to create an up-to-date list of items that have been in transit for longer than 5 days. It:
 *
 *   ** Queries the CARL reports server for items that are in transit and that meet the query criteria (see "Notes about
 *      this query").
 *   ** Parses through each result from the query and assigns it to each branch where it has had some contact.
 *   ** Clears out each existing sheet and dumps in new data. 
 */

var address = [* Your Reports IP/port *];
var username = 'reports';
var userPwd = 'carlx';
var db = [* Your Reports DB name *];
var dbUrl = 'jdbc:oracle:thin:@//' + address + '/' + db;

function getData() {
  /*
   * Notes about this query:
   *   ** This is pulls from a subquery made up of two queries smooshed together with a UNION statement (aliased as 'transit'). 
   *   ** The first part of the inner 'transit' query looks at both the status of the item (in transit for a hold) and the patron 
   *      attached to the hold from the transitem_v view, and looks for the most recent transit transaction date.
   *   ** The second part of the inner 'transit' query looks just at the status of the item (in transit, NOT for a hold) and looks 
   *      for the most recent transit transaction date.
   *   ** Note the whole MAX...OVER(PARTITION BY) thing to get the most recent transit transaction.
   *   ** The outer query gathers all of the information it needs to populate the Google Sheet, including the owning branch for
   *      the item.
   *   ** All string values for transaction type, status, etc. have their single-quotes escaped with a backslash. 
   */
 
  var sql = 'SELECT UNIQUE transit.item, b.title, i.cn, l.locname, transit.transitdate, transit.envbranch as transitfrom, ' +
            '   branch.branchcode as transitto, br.branchcode as owningbranch, transit.patronid ' +
            'FROM bbibmap_v b, location_v l, branch_v branch, branch_v br, item_v i ' +
            '    RIGHT JOIN ' +
            '    (SELECT item, jts.todate(systemtimestamp) as transitdate, envbranch, renew, patronid ' +
            '    FROM ' +
            '        (SELECT i.item, t.systemtimestamp, t.envbranch, ti.renew, ti.patronid, ' +
            '            (MAX(t.systemtimestamp) OVER(PARTITION BY i.item)) AS maxtransitdate ' +
            '        FROM item_v i, txlog_v t, transitem_v ti ' +
            '        WHERE i.status IN (\'IH\') AND ti.transcode = \'IH\' ' +
            '        AND jts.todate(t.systemtimestamp) < sysdate - 5 AND jts.todate(i.statusdate) < sysdate - 5 ' +
            '        AND i.item = t.item AND i.item = ti.item ' +
            '        AND ti.renew > 0 AND t.envbranch IS NOT NULL ' +
            '        AND t.transactiontype IN (\'DC\', \'DN\', \'DS\')) ' +
            '    WHERE systemtimestamp = maxtransitdate ' +
            '    UNION ' +
            '    SELECT item, jts.todate(systemtimestamp) as transitdate, envbranch, branch, \'\' as patronid ' +
            '    FROM ' +
            '        (SELECT i.item, t.systemtimestamp, t.envbranch, i.branch, ' +
            '            (MAX(t.systemtimestamp) OVER(PARTITION BY i.item)) AS maxtransitdate ' +
            '        FROM item_v i, txlog_v t ' +
            '        WHERE i.status IN (\'I\', \'IT\') ' +
            '        AND t.transactiontype IN (\'DC\', \'DN\', \'DS\') ' +
            '        AND jts.todate(t.systemtimestamp) < sysdate - 5 AND jts.todate(i.statusdate) < sysdate - 5 ' +
            '        AND i.item = t.item) ' +
            '    WHERE systemtimestamp = maxtransitdate) transit ' +
            '    ON i.item = transit.item ' +
            'WHERE i.location = l.locnumber AND transit.renew = branch.branchnumber ' +
            'AND i.owningbranch = br.branchnumber AND i.bid = b.bid';

  var conn = Jdbc.getConnection(dbUrl, username, userPwd);
  var stmt = conn.createStatement();
  var results = stmt.executeQuery(sql);
  
  // Set up a separate array for each branch, so that each tab can be quickly populated with its own set of branch-specific data.
  var bbdata = [];
  var brdata = [];
  var hidata = [];
  var madata = [];
  var mjdata = [];
  var npdata = [];
  var pedata = [];
  var sodata = [];
  var wadata = [];
  var wtdata = [];
  var scdata = [];
  bbdata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  brdata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  hidata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  madata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  mjdata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  npdata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  pedata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  sodata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  wadata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  wtdata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  scdata.push(['Item', 'Title', 'Call Number', 'Location', 'Transit Date', 'Transit From', 'Transit To', 'Owning', 'On Hold?']);
  
  while (results.next()) {
    // For each result from the query, add it to the array for each branch it's involved with: in transit from, in transit to
    // and owning.
    var from = results.getString(6);
    var to = results.getString(7);
    var own = results.getString(8);
    if (from == 'BB') {
      from = "BBROOK";
    } else if (from == 'BR') {  
      from = "BRIDGE";
    } else if (from == 'HI') {  
      from = "HILLSB";
    } else if (from == 'MA') {  
      from = "MANVLE";
    } else if (from == 'MJ') {  
      from = "MJACOB";
    } else if (from == 'NP') {  
      from = "NPLAIN";
    } else if (from == 'PE') {  
      from = "PEAGLA";
    } else if (from == 'SO') {  
      from = "SOMERV";
    } else if (from == 'WA') {  
      from = "WARREN";
    } else if (from == 'WT') {  
      from = "WTCHNG";
    } else if (from == 'ON') {  
      from = "ONLINE";
    } else if (from == 'SC') {  
      from = "SCLSNJ";
    } else if (from == 'BG') {  
      from = "BGL-RS";
    } else if (from == 'WV') {  
      from = "WVL-RS";
    } else {  
      from = "SCLSNJ";
    }
    if (to.match(/BBROOK/i) || from.match(/BBROOK/i) || own.match(/BBROOK/i)) {
      bbdata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    } 
    if (to.match(/BRIDGE/i) || from.match(/BRIDGE/i) || own.match(/BRIDGE/i)) {  
      brdata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
    if (to.match(/HILLSB/i) || from.match(/HILLSB/i) || own.match(/HILLSB/i)) {  
      hidata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
    if (to.match(/MANVLE/i) || from.match(/MANVLE/i) || own.match(/MANVLE/i)) {  
      madata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
    if (to.match(/MJACOB/i) || from.match(/MJACOB/i) || own.match(/MJACOB/i)) {  
      mjdata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
    if (to.match(/NPLAIN/i) || from.match(/NPLAIN/i) || own.match(/NPLAIN/i)) {  
      npdata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
    if (to.match(/PEAGLA/i) || from.match(/PEAGLA/i) || own.match(/PEAGLA/i)) {  
      pedata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
    if (to.match(/SOMERV/i) || from.match(/SOMERV/i) || own.match(/SOMERV/i)) {  
      sodata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
    if (to.match(/WARREN/i) || from.match(/WARREN/i) || own.match(/WARREN/i)) {  
      wadata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
    if (to.match(/WTCHNG/i) || from.match(/WTCHNG/i) || own.match(/WTCHNG/i)) {  
      wtdata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
    // Catch-all for items in transit to/from or owned by non-branch branches.
    if (to.match(/ONLINE/i) || from.match(/ONLINE/i) || own.match(/ONLINE/i) || to.match(/WVL-RS/i) || from.match(/WVL-RS/i) || own.match(/WVL-RS/i) || 
        to.match(/BGL-RS/i) ||from.match(/BGL-RS/i) || own.match(/BGL-RS/i) || to.match(/SCLSNJ/i) || from.match(/SCLSNJ/i) || own.match(/SCLSNJ/i)) {  
      scdata.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), from, to, own, results.getString(9)]);
    }
  }
  
  /* For each sheet (tab) in the spreadsheet (file):
   *   - Clear the existing content
   *   - Take the name of the sheet (which happens to the be the branch code) and use it to find the right array
   *   - Dump in the array values
   *   - Reformat the sheet: freeze the first row, sort the data, format the columns for width, number format
   */
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    sheet.getDataRange().clearContent();
    var name = sheet.getSheetName();
    var data = eval(name.substr(0,2).toLowerCase() + 'data');
    var range = sheet.getRange(1, 1, data.length, 9).setValues(data);
    var numRows = sheet.getLastRow();
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, numRows, 9).sort([4, 3]);
    var colFormats = [[120, ''],[400, ''],[180, ''],[180, ''],[80, 'mm/dd/yyyy'],[80, ''],[80, ''],[80, ''],[120, '']];
    for (var col = 0; col < colFormats.length; col++) {
      var range = sheet.getRange(1, col + 1, numRows);    
      sheet.setColumnWidth(col + 1, colFormats[col][0]);               // sets column width
      if (colFormats[col][1]) {
        range.setNumberFormat(colFormats[col][1]);                     // sets number format, if specified
      }
    }
  }
}

