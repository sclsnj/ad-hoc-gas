
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


function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('Get New Data')
    .addItem('Get New Data', 'getData')
    .addToUi();
}




function getData() {
  /*
   * Notes about this query:
   *   ** This took a TREMENDOUS amount of tinkering and fine-tuning. Please feel free to use this as a starting point, but
   *      be aware that you will have to do a lot of testing to make it work with your data.
   *   ** The CircDay element returns a number corresponding to the day of the week (1-7) that helps make the formulas in
   *      the Google Sheet work correctly.
   *   ** The CircHour element pulls the time portion of the date and then returns just the two digits corresponding to the hour.
   *   ** The <yesterday> and <today> variables in the query ensure that when this is triggered daily, it returns 
   *      yesterday's data.
   *   ** Because we use this data to get a sense of how busy we are from a staffing perspective, we're excluding non-branches
   *      from the results.
   *   ** The outer query groups the interior query results by date, hour and branch and returns a count.
   */

   var sql = 'SELECT bibs.bid, bibs.author, bibs.title, bibs.isbn, format.formattext, holds.holdcount, realitems.realcount, ' + 
            '    (CASE WHEN (onorder.onordercount = 0 OR onorder.onordercount IS NULL) THEN pending.copycount ELSE onorder.onordercount END) as ordercount ' +
            'FROM bbibmap_v bibs ' +
            'LEFT JOIN ' +
            '    (SELECT bid, COUNT(bid) as holdcount ' +
            '    FROM transbid_v ' +
            '    WHERE transcode = \'R*\' ' +
            '    GROUP BY bid) holds ' +
            'ON bibs.bid = holds.bid ' +
            'LEFT JOIN ' +
            '    (SELECT bid, COUNT(bid) as realcount ' +
            '    FROM item_v ' +
            '    WHERE media <> 42 AND media <> 43  ' +
            '    AND (status IN (\'C\', \'CT\', \'H\', \'HT\', \'I\', \'IT\', \'IH\', \'S\', \'ST\') OR (status = \'SP\' AND kbra <> \'SC\')) ' +
            '    AND kbra <> \'BG\' AND kbra <> \'WV\' ' +
            '    GROUP BY bid) realitems ' +
            'ON bibs.bid = realitems.bid ' +
            'LEFT JOIN ' +
            '    (SELECT bid, SUM(pending) as onordercount ' +
            '    FROM ' +
            '        (SELECT bid, (copycount - unitsreceivedcount) as pending ' +
            '        FROM orderdetail_v ' +
            '        WHERE status IN (\'N\', \'O\', \'O$\', \'R\', \'RC\', \'RF\', \'RG\', \'RI\', \'RN\', \'RP\') ' +
            '        AND fund <> \'2017ALL\') ' +
            '    GROUP BY bid) onorder ' +
            'ON bibs.bid = onorder.bid ' +
            'LEFT JOIN ' +
            '    (SELECT o.bid, to_number(o.copycount) as copycount, jts.todate(o.statusdate) as receiveddate ' +
            '    FROM orderdetail_v o  ' +
            '    INNER JOIN ' +
            '        (SELECT UNIQUE i.bid, jts.todate((MAX(i.creationdate) OVER(PARTITION BY i.bid))) AS maxcreateddate ' +
            '        FROM item_v i) itemcreated ' +
            '    ON itemcreated.bid = o.bid ' +
            '    WHERE o.destination NOT IN (\'BGL-RS\', \'WVL-RS\') ' +
            '    AND o.status IN (\'R\', \'RF\', \'RC\') ' +
            '    AND itemcreated.maxcreateddate < jts.todate(o.statusdate)) pending ' +
            'ON bibs.bid = pending.bid ' +
            'LEFT JOIN ' +
            'formatterm_v format ' +
            'ON bibs.format = format.formattermid ' +
            'WHERE (holds.holdcount/(CASE WHEN ((CASE WHEN realitems.realcount IS NOT NULL THEN realitems.realcount ELSE 0 END) +  ' +
            '(CASE WHEN onorder.onordercount IS NOT NULL THEN onorder.onordercount ELSE 0 END) +  ' +
            '(CASE WHEN pending.copycount IS NOT NULL THEN pending.copycount ELSE 0 END)) > 0  ' +
            'THEN ((CASE WHEN realitems.realcount IS NOT NULL THEN realitems.realcount ELSE 0 END) +  ' +
            '(CASE WHEN onorder.onordercount IS NOT NULL THEN onorder.onordercount ELSE 0 END) +  ' +
            '(CASE WHEN pending.copycount IS NOT NULL THEN pending.copycount ELSE 0 END)) ELSE 1 END)) > 4 ' +
            'AND holds.holdcount IS NOT NULL';
  
  var conn = Jdbc.getConnection(dbUrl, username, userPwd);
  var stmt = conn.createStatement();
  var results = stmt.executeQuery(sql);
  var data = [];
  data.push(['BID', 'Author', 'Title', 'ISBN', 'Format', 'Holds', 'Items', 'On Order', 'Ratio']);
  while (results.next()) {
    var holds = results.getString(6);
    var items = results.getString(7);
    var order = results.getString(8);
    if (!items) {
      items = 0;
    } else {
      items = Number(items);
    }
    if (!order) {
      order = 0;
    } else {
      order = Number(order);
    }
    var div = items + order;
    if (div == 0) {
      div = 1;
    }
    var ratio = holds/div;
    var format = results.getString(5);
    if (format) {
      if (format.match(/DVD/i) || format.match(/Blu/i)) {
        if (ratio >= 10) {
          data.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), format, holds, items, order, ratio]);
        }
      } else if (format.match(/book$/i) || format.match(/large/i)) {
        if (ratio >= 4) {
          data.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), format, holds, items, order, ratio]);
        }
      } else if (format.match(/music/i)) {
        if (ratio >= 5) {
          data.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), format, holds, items, order, ratio]);
        }
      } else if (format.match(/CD/i) || format.match(/playaway/i)) {
        if (ratio >= 6) {
          data.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), format, holds, items, order, ratio]);
        }
      }  
    } else {
      data.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), format, holds, items, order, ratio]);
    }
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var date = new Date();
  var name = date.toLocaleDateString();
  var exists = ss.getSheetByName(name);
  if (exists) {
    name += ' (2)'
  }
  var sheet = ss.insertSheet(name);
  var range = sheet.getRange(1, 1, data.length, 9).setValues(data);
  var numRows = sheet.getLastRow();
  sheet.setFrozenRows(1);
  sheet.getRange(2, 1, numRows, 9).sort([5, 9]);
  var colFormats = [[81, '0'],[240, ''],[452, ''],[100, ''],[94, ''],[75, '0'],[75, '0'],[75, '0'],[75, '0']];
  for (var col = 0; col < colFormats.length; col++) {
    var range = sheet.getRange(1, col + 1, numRows);    
    sheet.setColumnWidth(col + 1, colFormats[col][0]);               // sets column width
    if (colFormats[col][1]) {
      range.setNumberFormat(colFormats[col][1]);                     // sets number format, if specified
    }
  }
}
