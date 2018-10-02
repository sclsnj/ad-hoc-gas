function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('Get New Data')
    .addItem('Get New Data', 'getData')
    .addToUi();
}


// Ad Hoc Database credentials
var address = '209.212.22.12:1521';
var username = 'reports';
var userPwd = 'carlx';
var db = 'somprod';
var dbUrl = 'jdbc:oracle:thin:@//' + address + '/' + db;


function getData() {

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
