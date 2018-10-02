

function onOpen(e) {
  var conn = Jdbc.getConnection(dbUrl, username, userPwd);
  var stmt = conn.createStatement(); 
  var sql = 'SELECT v.vendorcode, encumbered.List, expended.Spent ' +
            'FROM vendor_v v ' +
            'FULL JOIN ' +
            '    (SELECT v1.vendorcode, SUM(o.total) AS List ' +
            '    FROM vendor_v v1, ' +
            '        (SELECT vendorcode, price*copycount AS total ' +
            '         FROM orderdetail_v ' +
            '         WHERE fund LIKE \'2018%\' AND (status = \'O\' OR status = \'N\')) o ' +
            '    WHERE v1.vendorcode = o.vendorcode ' +
            '    GROUP BY v1.vendorcode) encumbered ' +
            'ON v.vendorcode = encumbered.vendorcode ' +
            'FULL JOIN ' +
            '    (SELECT v2.vendorcode, SUM(r.total) AS Spent ' +
            '    FROM vendor_v v2, ' +
            '        (SELECT vendorcode, cost*copycount AS total ' +
            '         FROM orderdetail_v ' +
            '         WHERE fund LIKE \'2018%\' AND (status <> \'O\' AND status <> \'N\')) r ' +
            '    WHERE v2.vendorcode = r.vendorcode ' +
            '    GROUP BY v2.vendorcode) expended ' +
            'ON v.vendorcode = expended.vendorcode ' +
            'WHERE encumbered.List IS NOT NULL or expended.Spent IS NOT NULL ' +
            'ORDER BY v.vendorcode';
  var results = stmt.executeQuery(sql);
  var vendors = [];
  while (results.next()) {
    vendors.push([results.getString(1), results.getString(2), results.getString(3)]);
  }
  for (var v = 0; v < vendors.length; v++) {
    var vendorcode = vendors[v][0];
    if (vendorcode.match(/^BT/)) {
      vendors[v][0] = 'BT';
    } else if (vendorcode.match(/^MID/)) {
      vendors[v][0] = 'MID';
    } else if (vendorcode.match(/^THP/)) {
      vendors[v][0] = 'GAL';
    }
  }
  var totals = [];
  for (var v = 0; v < vendors.length; v++) {
    var currvendor = vendors[v][0];
    var duplicate = false;
    total : for (var t = 0; t < totals.length; t++) {
      var testvendor = totals[t][0];
      if (currvendor == testvendor) {
        duplicate = true;        
        totals[t][1] = totals[t][1] + Number(vendors[v][1]);
        totals[t][2] = totals[t][2] + Number(vendors[v][2]);
        break total;
      }
    }
    if (duplicate == false) {
      totals.push([currvendor, Number(vendors[v][1]) ,Number(vendors[v][2])]);
    }
  }
  var existingdata = [];
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('2018 Vendor Spending');
  var lastRow = ss.getLastRow();
  var range = sheet.getRange(3, 1, lastRow - 2, 9);
  existingdata = range.getValues();
  for (var d = 0; d < existingdata.length; d++) {
    var testvendor = existingdata[d][0];
    vend : for (var t = 0; t < totals.length; t++) {
      var vendorcode = totals[t][0];
      if (vendorcode == testvendor) {
        var encumber = totals[t][1];
        var spend = totals[t][2];
        if (encumber != null) {
          existingdata[d][8] = encumber;
        }
        if (spend != null) {
          existingdata[d][6] = spend;
        }
        break vend;
      }
    }
    existingdata[d][7] = '=(C' + (d+3) + '+D' + (d+3) + '+E' + (d+3) + '+F' + (d+3) + ')-G' + (d+3);
  }
  range.setValues(existingdata);
  var today = new Date();
  today = Utilities.formatDate(today, "GMT-0500", "MM-dd-yyyy");
  range = sheet.getRange(1, 5);
  range.setValue(today);
}
