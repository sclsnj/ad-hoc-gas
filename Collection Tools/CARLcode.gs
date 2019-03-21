// Called from the sidebar on submitting the form, this function constructs the ad hoc query, 
// receives and handles data, and pushes it into the Google Sheet.
function formatCARLCollCheck(formData) {
  var createdstart, createdend, statusstart, statusend, modifiedstart, modifiedend, status, branch, location, media;
  var useOldandout, filterOldandout, pubdateAge, useNocirc, filterNocirc, ladAge, useOldwithrecent, filterOldwithrecent, createdAge; 
  var useSupercirc, filterSupercirc, bigCirc, useConstantcirc, filterConstantcirc, circRate;
  createdstart = formData.createdstart;
  createdend = formData.createdend;
  statusstart = formData.statusstart;
  statusend = formData.statusend;
  modifiedstart = formData.modifiedstart;
  modifiedend = formData.modifiedend;
  status = formData.status;
  branch = formData.branch;
  location = formData.location;
  media = formData.media;
  useOldandout = formData.useoldandout;
  filterOldandout = formData.filteroldandout;
  pubdateAge = formData.pubdateage;
  useNocirc = formData.usenocirc;
  filterNocirc = formData.filternocirc;
  ladAge = formData.ladage;
  useOldwithrecent = formData.useoldwithrecent;
  filterOldwithrecent = formData.filteroldwithrecent;
  createdAge = formData.createdage;
  useSupercirc = formData.usesupercirc;
  filterSupercirc = formData.filtersupercirc;
  bigCirc = formData.bigcirc;
  useConstantcirc = formData.useconstantcirc;
  filterConstantcirc = formData.filterconstantcirc;
  circRate = formData.circrate;
 
  var saveprefs = formData.saveprefs;
  // Save current settings as default for this user, if requested
  if (saveprefs == true) {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('useOldandout', useOldandout);
    userProperties.setProperty('pubdateAge', pubdateAge);
    userProperties.setProperty('useNocirc', useNocirc);
    userProperties.setProperty('ladAge', ladAge);
    userProperties.setProperty('useOldwithrecent', useOldwithrecent);
    userProperties.setProperty('createdAge', createdAge);
    userProperties.setProperty('useSupercirc', useSupercirc);
    userProperties.setProperty('bigCirc', bigCirc);
    userProperties.setProperty('useConstantcirc', useConstantcirc);
    userProperties.setProperty('circRate', circRate);
  }
  
  // The comment variable becomes a sort-of title for the final sheet; it could be prettier...
  var comment = 'Collection Check Parameters: \n';
  
  // The base query.
  var sql = 'SELECT UNIQUE i.bid, i.item, b.author, b.title, b.publishingdate, i.cn, status.description, i.price, branch.branchcode, location.locname, ' +
            '  media.medcode, i.circhistory, i.cumulativehistory, (CASE WHEN t.circdate IS NOT NULL THEN t.circdate ELSE jts.todate(i.statusdate) END), ' +
            '  (CASE WHEN b.isbn IS NOT NULL THEN b.isbn ELSE (CASE WHEN b.upc IS NOT NULL THEN b.upc ELSE null END) END) as ISBN, ' +
            '  jts.todate(i.creationdate) as Created, jts.todate(i.editdate) as Edited, jts.todate(i.statusdate) as LastActivity, i.inhousecirc ' +
            'FROM item_v i ' +
            'JOIN bbibmap_v b ' +
            'ON i.bid = b.bid ' +
            'LEFT JOIN ' +
            '    (SELECT item, jts.todate(systemtimestamp) as circdate ' +
            '    FROM ' +
            '        (SELECT item, systemtimestamp, (MAX(systemtimestamp) OVER(PARTITION BY item)) AS maxcircdate ' +
            '        FROM txlog_v ' +
            '        WHERE transactiontype = \'CH\') ' +
            '    WHERE systemtimestamp = maxcircdate) t '+
            'ON i.item = t.item ' +
            'JOIN systemitemcodes_v status ' +
            'ON i.status = status.code ' +
            'JOIN branch_v branch ' +
            'ON i.branch = branch.branchnumber ' +
            'JOIN location_v location ' +
            'ON i.location = location.locnumber ' +
            'JOIN media_v media ' +
            'ON i.media = media.mednumber ' +
            'WHERE b.bibtype = 0 ' +
            'AND b.acqtype = 0 ';
  
  // These conditions add to the query if any of the date limits are selected.
  // NOTE: There's probably a better way to do this date formatting than using the formatDateForSQL function down below,
  //       but dates are the bane of my existance, and this got the job done. -LH
  if (createdstart) {
    sql += 'AND jts.todate(i.creationdate) > \'' + formatDateForSql(createdstart) + '\' ';
    comment += '- created on or after ' + createdstart + '\n';
  }
  if (createdend) {
    sql += 'AND jts.todate(i.creationdate) < \'' + formatDateForSql(createdend) + '\' ';
    comment += '- created before ' + createdend + '\n';
  }
  if (statusstart) {
    sql += 'AND jts.todate(i.statusdate) > \'' + formatDateForSql(statusstart) + '\' ';
    comment += '- status updated on or after ' + statusstart + '\n';
  }
  if (statusend) {
    sql += 'AND jts.todate(i.statusdate) < \'' + formatDateForSql(statusend) + '\' ';
    comment += '- status updated before ' + statusend + '\n';
  }
  if (modifiedstart) {
    sql += 'AND jts.todate(i.editdate) > \'' + formatDateForSql(modifiedstart) + '\' ';
    comment += '- item edited on or after ' + modifiedstart + '\n';
  }
  if (modifiedend) {
    sql += 'AND jts.todate(i.editdate) < \'' + formatDateForSql(modifiedend) + '\' ';
    comment += '- item edited before ' + modifiedend + '\n';
  }
  
  // This ensures that the status part of the query behaves appropriately whether one or multiple statuses are selected.
  if (status) {
    if (Array.isArray(status)) {
      sql += 'AND i.status IN (';
      for (var s = 0; s < status.length - 1; s++) {
        sql += '\'' + status[s] + '\',';
      }
      sql += '\'' + status[status.length - 1] + '\') ';
    } else {
      sql += 'AND i.status IN (' + status + ') ';
    }
    comment += '- status(es): ' + status + '\n';
  }
  
  // This ensures that the branch part of the query behaves appropriately whether one or multiple branches are selected.
  if (branch) {
    if (Array.isArray(branch)) {
      sql += 'AND i.branch IN (';
      for (var b = 0; b < branch.length - 1; b++) {
        sql += branch[b] + ',';
      }
      sql += branch[branch.length - 1] + ') ';
    } else {
      sql += 'AND i.branch IN (' + branch + ') ';
    }
    comment += '- branch(es): ' + branch + '\n';
  }
  
  // This ensures that the location part of the query behaves appropriately whether one or multiple locations are selected.
  if (location) {
    if (Array.isArray(location)) {
      sql += 'AND i.location IN (';
      for (var l = 0; l < location.length - 1; l++) {
        sql += location[l] + ',';
      }
      sql += location[location.length - 1] + ') ';
    } else {
      sql += 'AND i.location IN (' + location + ') ';
    }
    comment += '- location(s): ' + location + '\n';
  }
  
  // This ensures that the media part of the query behaves appropriately whether one or multiple media types are selected.
  if (media) {
    if (Array.isArray(media)) {
      sql += 'AND i.media IN (';
      for (var m = 0; m < media.length - 1; m++) {
        sql += media[m] + ',';
      }
      sql += media[media.length - 1] + ') ';
    } else {
      sql += 'AND i.media IN (' + media + ') ';
    }
    comment += '- media type(s): ' + media + '\n';
  }
 
  if (filterOldandout) {
    sql += 'AND (sysdate - (SELECT TO_DATE((\'01-01-\'||TO_NUMBER(b.publishingdate)),\'MM-DD-YYYY\') FROM DUAL))/365 >= ' + pubdateAge + ' ' +
           'AND i.status IN (\'C\', \'CT\', \'H\', \'IH\') ';
    comment += '- currently checked out with a publication date ' + pubdateAge + ' years old or older\n';
  }
  if (filterNocirc) {
    sql += 'AND (sysdate - jts.todate(i.statusdate))/30 < ' + ladAge + ' ';
    comment += '- no recent activity within the last ' + pubdateAge + ' months\n';
  }
  if (filterOldwithrecent) {
    sql += 'AND (sysdate - jts.todate(i.creationdate))/365 >= ' + createdAge + ' ';
    comment += '- recent activity but that we\'ve had for ' + pubdateAge + ' years or longer\n';
  }
  if (filterSupercirc) {
    sql += 'AND i.cumulativehistory >= ' + bigCirc + ' ';
    comment += '- at least  ' + pubdateAge + ' circs\n';
  }
  if (filterConstantcirc) {
    sql += 'AND (i.cumulativehistory/((sysdate - jts.todate(i.creationdate))/365)) >= ' + circRate + ' ';
    comment += '- circulated at least  ' + pubdateAge + ' times per year since we\'ve had them\n';
  }
  
  // You can, of course, store these as script parameters and pull from there rather than including in the code.
  var address = [* Your Reports IP/port *];
  var username = 'reports';
  var userPwd = 'carlx';
  var db = [* Your Reports DB name *];
  
  var dbUrl = 'jdbc:oracle:thin:@//' + address + '/' + db;
  var conn = Jdbc.getConnection(dbUrl, username, userPwd);
  var stmt = conn.createStatement();
  var results = stmt.executeQuery(sql);
  var data = []; 
  if (results) {
    while (results.next()) {
      data.push([results.getString(1), results.getString(2), results.getString(3), results.getString(4), results.getString(5), results.getString(6), 
                results.getString(7), results.getString(8), results.getString(9), results.getString(10), results.getString(11), results.getString(12), 
                results.getString(13), results.getString(14), results.getString(15), results.getString(16), 
                results.getString(17), results.getString(18), results.getString(19), '']);
    }
  } else {
    Logger.log(results.getWarnings());
  }
  conn.close();
  
  var today = new Date();
  for (var x = 0; x < data.length; x++) {
    var flag = '';
    
    // Standardizes pubdate as an actual date, for better date comparison and sorting; it also arbitrarily makes the pub date
    // ten years ago if it can't make sense of the pub year coming from the query.
    if (data[x][5]) {
      var pubyear = data[x][5].toString();
    } else {
      var pubyear = today.getFullYear - 10;
    }
    var pubdate = new Date('01/01/' + pubyear);
    
    // Flags! This whole section populates the added "Flag" column in the spreadsheet, which we use to give branch staff some
    // added info about the items the might be searching for. 
    // NOTE: The order in which these are processed is based on our rules for which flag should show if more than one applies...
    var circs = Number(data[x][12]);
    if (data[x][13]) {
      var lastCirc = new Date(data[x][13].substr(0,10));
    }
    var creationDate = new Date(data[x][15].substr(0,10));
    var age = (today.getTime() - creationDate.getTime())/(1000 * 60 * 60 * 24 * 365);
    if (useNocirc) {
      if (lastCirc) {
        if (new Date(lastCirc.getTime() + (ladAge/12) * 365 * 24 * 60 * 60 * 1000) < today) {
          flag = 'No recent activity - deselect?';
        }
      } else {
        if ((new Date(creationDate.getTime() + (ladAge/12) * 365 * 24 * 60 * 60 * 1000) < today) && (circs < 1)) {
          flag = 'Zero circs, ' + ladAge + '+ months old';
        }
        else if ((new Date(creationDate.getTime() + (ladAge/12) * 365 * 24 * 60 * 60 * 1000) < today) && (circs >= 1)) {
          flag = 'No recent activity - deselect?';
        }
      }
    }
    if (useOldwithrecent) {
      if ((new Date(lastCirc.getTime() + ((ladAge/12) * 365 * 24 * 60 * 60 * 1000)) >= today) && 
          (new Date(creationDate.getTime() + createdAge * 365 * 24 * 60 * 60 * 1000) < today)) {
        flag = createdAge + '+ years old - deselect or replace?';
      }
    }    
    if (useOldandout) {
      if (pubdate) {
        if (data[x][6].match(/checked out/i) && (new Date(pubdate.getTime() + pubdateAge * 365 * 24 * 60 * 60 * 1000) < today)) {
          flag = 'Checked out, older pub - replace?';
        }
      }
    }
    if (useConstantcirc) {
      var avgCirc = circs/age;
      if ((avgCirc > circRate)) {
        flag = circRate + '+ circs per year - replace?';
      }
    }
    if (useSupercirc) {
      if (circs > bigCirc) {
      flag = bigCirc + '+ circs - deselect or replace?';
      }
    }
    if (flag) {
      data[x][18] = flag;
    }
    flag = '';
  }
 
  // Put all of the processed query data in the spreadsheet, starting with row 2 so that we can add the header in row 1
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.getRange(2, 1, data.length, 20).setValues(data);
  // Add header row
  var headers = [['BID', 'Item', 'Author', 'Title', 'Pub Date', 'Call #', 'Status', 'Price', 'Branch', 'Location', 'Media', 
                 'Current Circ', 'Total Circ', 'Last Circ', 'ISBN', 'Created Date', 'Edited Date', 'Status Date', 'In House Use', 'Flag']];
  sheet.getRange(1, 1, 1, 20).setValues(headers);
  
  // Calls the formatReport function from code.gs, which is what makes number formats, conditional formatting, etc. work
  formatReport(sheet);

  // If the comment variable from up above has something in it, stick it in an inserted row at the top of the sheet.
  if (comment) {
    sheet.insertRowBefore(1);
    sheet.getRange(1,1).setValue(comment);
  }
}


function formatDateForSql(date) {
  var properDate = new Date(date);
  var day = properDate.getUTCDate();
  if (day < 10) {
    day = '0' + day;
  }
  var year = properDate.getUTCFullYear().toString();
  year = year.substr(2,2);
  var month = properDate.getUTCMonth();
  if (month == 0) {
    return day + '-JAN-' + year;
  } else if (month == 1) {
    return day + '-FEB-' + year;
  } else if (month == 2) {
    return day + '-MAR-' + year;
  } else if (month == 3) {
    return day + '-APR-' + year;
  } else if (month == 4) {
    return day + '-MAY-' + year;
  } else if (month == 5) {
    return day + '-JUN-' + year;
  } else if (month == 6) {
    return day + '-JUL-' + year;
  } else if (month == 7) {
    return day + '-AUG-' + year;
  } else if (month == 8) {
    return day + '-SEP-' + year;
  } else if (month == 9) {
    return day + '-OCT-' + year;
  } else if (month == 0) {
    return day + '-NOV-' + year;
  } else if (month == 11) {
    return day + '-DEC-' + year;
  }
}
