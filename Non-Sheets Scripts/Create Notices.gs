function myFunction() {
  var address = '209.212.22.12:1521';
  var username = 'reports';
  var userPwd = 'carlx';
  var db = 'somprod';
  var dbUrl = 'jdbc:oracle:thin:@//' + address + '/' + db;

  // Get notice data
  var today = new Date();
  var day = today.getDay();
  if (day == 1) {
    var daysback = 3;
  } else {
    var daysback = 1;
  }    
  var conn = Jdbc.getConnection(dbUrl, username, userPwd);
  var stmt = conn.createStatement();
  var sql = 'SELECT n.item, n.patronid, p.firstname, p.middlename, p.lastname, p.suffixname, p.street1, p.city1, p.state1, p.zip1, ' +
            '  jts.todate(n.transdate) as transactiondate, jts.todate(n.dueornotneededafterdate) as nnadate, ' +
            '  jts.todate(n.lastactiondate) as actiondate, jts.todate(n.returndate) as returned, ' +
            '  br.branchname as holdingbranch, br.branchaddress1, br.branchcity, br.branchzipcode, m.medname, ' +
            '  l.locname, n.amountdebited, n.noticetype, n.noticeformat, n.noticedeliverytype, b.title, b.author, b.callnumber, ' +
            '  n.noticedate, nbr.branchname, nbr.branchaddress1, nbr.branchcity, nbr.branchzipcode, nbr.branchphone ' +
            'FROM notices_v n, patron_v p, media_v m, location_v l, bbibmap_v b, item_v i, ' +
            '  (SELECT branchname, branchaddress1, branchcity, branchzipcode, branchnumber FROM branch_v) br, ' +
            '  (SELECT branchname, branchaddress1, branchcity, branchzipcode, branchnumber, branchphone FROM branch_v) nbr ' +
            'WHERE n.patronid = p.patronid ' +
            'AND n.media = m.mednumber ' +
            'AND n.location = l.locnumber ' +
            'AND n.item = i.item ' +
            'AND i.bid = b.bid ' +
            'AND n.noticebranch = nbr.branchnumber ' +
            'AND n.holdingbranch = br.branchnumber ' +
            'AND n.noticedeliverytype = \'P\' ' +
            'AND n.processed = \'N\' ' +
            'AND p.bty NOT IN (12, 14) ' +
            'AND n.noticedate >= sysdate - ' + daysback + ' ' +
            'ORDER BY n.noticetype, n.patronid';
  var results = stmt.executeQuery(sql);

  var numCols = results.getMetaData().getColumnCount();

  var holdNoticeItems = [];
  var lostNoticeItems = [];
  var overdueNoticeItems = [];
  while (results.next()) {
    var row = [];
    for (var col = 0; col < numCols; col++) {
      row.push(results.getString(col + 1));
    }
    var noticeType = row[21];
    if (noticeType == 'L1') {
      lostNoticeItems.push(row);
    } else if (noticeType == 'HL') {
      holdNoticeItems.push(row);
    } else if (noticeType == 'O2') {
      overdueNoticeItems.push(row);
    }
  }
  Logger.log(overdueNoticeItems);
  // Create today's notice file
  var date = Utilities.formatDate(new Date(), "GMT-0500", "YYYY MM dd");
  var doc = DocumentApp.create('Notices - ' + date);
  var body = doc.getBody();
  var docId = doc.getId();

  // Define custom paragraph styles
  var systemstyle = {};
  systemstyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  systemstyle[DocumentApp.Attribute.FONT_FAMILY] = 'Open Sans';
  systemstyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  systemstyle[DocumentApp.Attribute.BOLD] = false;

  var branchstyle = {};
  branchstyle[DocumentApp.Attribute.SPACING_AFTER] = 18;
  branchstyle[DocumentApp.Attribute.BOLD] = false;
  
  var itemstyle = {};
  itemstyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  itemstyle[DocumentApp.Attribute.FONT_FAMILY] = 'Open Sans';
  itemstyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  itemstyle[DocumentApp.Attribute.INDENT_START] = 30;
  itemstyle[DocumentApp.Attribute.INDENT_FIRST_LINE] = 30;
  itemstyle[DocumentApp.Attribute.BOLD] = false;

  var alertstyle = {};
  alertstyle[DocumentApp.Attribute.BOLD] = true;
  alertstyle[DocumentApp.Attribute.SPACING_AFTER] = 18;
  
  // Loop through lost/bill notices by user
  var newPatron = true;
  for (var i = 0; i < lostNoticeItems.length; i++) {
    var noticetype = lostNoticeItems[i][21];
    if (newPatron) {
      var system = body.appendParagraph('Somerset County Library System of New Jersey');
      system.setAttributes(systemstyle);
      var branchAddress = lostNoticeItems[i][29];
      if (!branchAddress) {
        var branch = body.appendParagraph('1 Vogt Dr\r' +
                                          'Bridgewater NJ 08807 \r\r');
        var contact = 'Somerset County Library System at 908-458-8400';
      } else {
        var branch = body.appendParagraph(lostNoticeItems[i][28] + '\r' +
                                        lostNoticeItems[i][29] + '\r' +
                                        lostNoticeItems[i][30] + ' ' + lostNoticeItems[i][31] + '\r');
        branch.setAttributes(branchstyle);
        var contact = lostNoticeItems[i][28] + ' at ' + lostNoticeItems[i][32];
      }
      branch.setAttributes(branchstyle);
      var address = body.appendParagraph(lostNoticeItems[i][2] + ' ' + lostNoticeItems[i][4] + '\r' +
                                        lostNoticeItems[i][6] + '\r' +
                                        lostNoticeItems[i][7] + ' ' + lostNoticeItems[i][8] + ' ' + lostNoticeItems[i][9] + '\r');
      address.setAttributes(branchstyle);
      var d = lostNoticeItems[i][27];
      var noticeDate = Utilities.formatDate(new Date(d.substr(0,4) + '/' + d.substr(5,2) + '/' + d.substr(8,8)), "GMT-0500", "MMMM d, YYYY");
      var intro = body.appendParagraph(noticeDate);
      intro.setAttributes(branchstyle);
      var name = lostNoticeItems[i][2];
      if (!name) {
        name = lostNoticeItems[i][4];
      }
      var properName = titleCase(name);
      var open = body.appendParagraph('Dear ' + properName + ',');
      open.setAttributes(branchstyle);
      var notice = body.appendParagraph('BILLING NOTICE');
      notice.setAttributes(alertstyle);
      var para = body.appendParagraph('Your account has been blocked and you are being billed the replacement cost for the library material below, which is checked out on your library card and is more than 5 weeks overdue. Previous notices were sent regarding this material by email, phone or postal mail.');
      para.setAttributes(branchstyle);
    }
    var titleText = lostNoticeItems[i][24];
    var author = lostNoticeItems[i][25];
    if (author) {
      titleText += ' - ' + author;
    }
    var title = body.appendParagraph(titleText);
    title.setAttributes(itemstyle);
    var call = body.appendParagraph(lostNoticeItems[i][26] + ' (' + lostNoticeItems[i][0] + ')');
    call.setAttributes(itemstyle);
    var d = lostNoticeItems[i][11];
    var dueDate = Utilities.formatDate(new Date(d.substr(0,4) + '/' + d.substr(5,2) + '/' + d.substr(8,8)), "GMT-0500", "EEEE, MMMM d, YYYY");
    var due = body.appendParagraph('This item was due ' + dueDate);
    due.setAttributes(itemstyle);
    var price = lostNoticeItems[i][20];
    if (price.match(/00$/)) {
      var cost = body.appendParagraph('REPLACEMENT COST: $' + (lostNoticeItems[i][20])/100 + '.00');
    } else if (price.match(/0$/)) {
      var cost = body.appendParagraph('REPLACEMENT COST: $' + (lostNoticeItems[i][20])/100 + '0');
    } else {
      var cost = body.appendParagraph('REPLACEMENT COST: $' + (lostNoticeItems[i][20])/100);
    }
    cost.setAttributes(itemstyle);
    cost.setAttributes(alertstyle);
    
    var patron = lostNoticeItems[i][1];
    if (i + 1 != lostNoticeItems.length) {
      var nextPatron = lostNoticeItems[i+1][1];
      if (patron == nextPatron) {
        newPatron = false;
      } else {
        var summary = body.appendParagraph('Payment is due immediately. Account balances of $75.00 or more will be submitted to the collection agency in 7 days.');
        summary.setAttributes(branchstyle);
        var closing = body.appendParagraph('For more information, visit SCLSNJ.org to view your account online. If you have any questions, you can call the ' +
                                           contact + ', or email ask@sclibnj.org.');
        closing.setAttributes(branchstyle);
        body.appendPageBreak();
        newPatron = true;
      }
    } else {
      var summary = body.appendParagraph('Payment is due immediately. Account balances of $75.00 or more will be submitted to the collection agency in 7 days.');
      summary.setAttributes(branchstyle);
      var closing = body.appendParagraph('For more information, visit SCLSNJ.org to view your account online. If you have any questions, you can call the ' +
                                           contact + ', or email ask@sclibnj.org.');
      closing.setAttributes(branchstyle);
      body.appendPageBreak();
    }
  }

  // Loop through hold notices by user
  var newPatron = true;
  for (var i = 0; i < holdNoticeItems.length; i++) {
    var noticetype = holdNoticeItems[i][21];
    if (newPatron) {
      var system = body.appendParagraph('Somerset County Library System of New Jersey');
      system.setAttributes(systemstyle);
      var branchAddress = holdNoticeItems[i][29];
      if (!branchAddress) {
        var branch = body.appendParagraph('1 Vogt Dr\r' +
                                          'Bridgewater NJ 08807 \r\r');
        var contact = 'Somerset County Library System at 908-458-8400';
      } else {
        var branch = body.appendParagraph(holdNoticeItems[i][28] + '\r' +
                                        holdNoticeItems[i][29] + '\r' +
                                        holdNoticeItems[i][30] + ' ' + holdNoticeItems[i][31] + '\r');
        branch.setAttributes(branchstyle);
        var contact = holdNoticeItems[i][28] + ' at ' + holdNoticeItems[i][32];
      }
      branch.setAttributes(branchstyle);
      var address = body.appendParagraph(holdNoticeItems[i][2] + ' ' + holdNoticeItems[i][4] + '\r' +
                                        holdNoticeItems[i][6] + '\r' +
                                        holdNoticeItems[i][7] + ' ' + holdNoticeItems[i][8] + ' ' + holdNoticeItems[i][9] + '\r');
      address.setAttributes(branchstyle);
      var d = holdNoticeItems[i][27];
      var noticeDate = Utilities.formatDate(new Date(d.substr(0,4) + '/' + d.substr(5,2) + '/' + d.substr(8,8)), "GMT-0500", "MMMM d, YYYY");
      var intro = body.appendParagraph(noticeDate);
      intro.setAttributes(branchstyle);
      var name = holdNoticeItems[i][2];
      if (!name) {
        name = holdNoticeItems[i][4];
      }
      var properName = titleCase(name);
      var open = body.appendParagraph('Dear ' + properName + ',');
      open.setAttributes(branchstyle);
      var notice = body.appendParagraph('Hold Notice');
      notice.setAttributes(alertstyle);
      var para = body.appendParagraph('The following material you placed on hold is ready for pickup:');
      para.setAttributes(branchstyle);
    }
    var titleText = holdNoticeItems[i][24];
    var author = holdNoticeItems[i][25];
    if (author) {
      titleText += ' - ' + author;
    }
    var title = body.appendParagraph(titleText);
    title.setAttributes(itemstyle);
    var call = body.appendParagraph(holdNoticeItems[i][26]);
    call.setAttributes(itemstyle);
    var d = holdNoticeItems[i][11];
    var pickupDate = Utilities.formatDate(new Date(d.substr(0,4) + '/' + d.substr(5,2) + '/' + d.substr(8,8)), "GMT-0500", "EEEE, MMMM d, YYYY");
    var pickup = body.appendParagraph('Ready at the ' + holdNoticeItems[i][28] + ' through ' + pickupDate);
    pickup.setAttributes(itemstyle);
    pickup.setSpacingAfter(12);
    
    var patron = holdNoticeItems[i][1];
    if (i + 1 != holdNoticeItems.length) {
      var nextPatron = holdNoticeItems[i+1][1];
      if (patron == nextPatron) {
        newPatron = false;
      } else {
        var closing = body.appendParagraph('For more information, visit SCLSNJ.org to view your account online. If you have any questions, you can call the ' +
                                           contact + ', or email ask@sclibnj.org.');
        closing.setAttributes(branchstyle);
        body.appendPageBreak();
        newPatron = true;
      }
    } else {
      var closing = body.appendParagraph('For more information, visit SCLSNJ.org to view your account online. If you have any questions, you can call the ' +
                                           contact + ', or email ask@sclibnj.org.');
      closing.setAttributes(branchstyle);
      body.appendPageBreak();
    }
  }

  
  // Loop through overdue notices by user
  var newPatron = true;
  for (var i = 0; i < overdueNoticeItems.length; i++) {
    var noticetype = overdueNoticeItems[i][21];
    if (newPatron) {
      var system = body.appendParagraph('Somerset County Library System of New Jersey');
      system.setAttributes(systemstyle);
      var branchAddress = overdueNoticeItems[i][29];
      if (!branchAddress) {
        var branch = body.appendParagraph('1 Vogt Dr\r' +
                                          'Bridgewater NJ 08807 \r\r');
        var contact = 'Somerset County Library System at 908-458-8400';
      } else {
        var branch = body.appendParagraph(overdueNoticeItems[i][28] + '\r' +
                                        overdueNoticeItems[i][29] + '\r' +
                                        overdueNoticeItems[i][30] + ' ' + overdueNoticeItems[i][31] + '\r');
        branch.setAttributes(branchstyle);
        var contact = overdueNoticeItems[i][28] + ' at ' + overdueNoticeItems[i][32];
      }
      branch.setAttributes(branchstyle);
      var address = body.appendParagraph(overdueNoticeItems[i][2] + ' ' + overdueNoticeItems[i][4] + '\r' +
                                        overdueNoticeItems[i][6] + '\r' +
                                        overdueNoticeItems[i][7] + ' ' + overdueNoticeItems[i][8] + ' ' + overdueNoticeItems[i][9] + '\r');
      address.setAttributes(branchstyle);
      var d = overdueNoticeItems[i][27];
      var noticeDate = Utilities.formatDate(new Date(d.substr(0,4) + '/' + d.substr(5,2) + '/' + d.substr(8,8)), "GMT-0500", "MMMM d, YYYY");
      var intro = body.appendParagraph(noticeDate);
      intro.setAttributes(branchstyle);
      var name = overdueNoticeItems[i][2];
      if (!name) {
        name = overdueNoticeItems[i][4];
      }
      var properName = titleCase(name);
      var open = body.appendParagraph('Dear ' + properName + ',');
      open.setAttributes(branchstyle);
      var notice = body.appendParagraph('OVERDUE NOTICE');
      notice.setAttributes(alertstyle);
      var para = body.appendParagraph('The following material checked out on your library card is currently overdue.');
      para.setAttributes(branchstyle);
    }
    var titleText = overdueNoticeItems[i][24];
    var author = overdueNoticeItems[i][25];
    if (author) {
      titleText += ' - ' + author;
    }
    var title = body.appendParagraph(titleText);
    title.setAttributes(itemstyle);
    var call = body.appendParagraph(overdueNoticeItems[i][26] + ' (' + overdueNoticeItems[i][0] + ')');
    call.setAttributes(itemstyle);
    var d = overdueNoticeItems[i][11];
    var dueDate = Utilities.formatDate(new Date(d.substr(0,4) + '/' + d.substr(5,2) + '/' + d.substr(8,8)), "GMT-0500", "EEEE, MMMM d, YYYY");
    var due = body.appendParagraph('This item was due ' + dueDate);
    due.setAttributes(itemstyle);
    due.setAttributes(alertstyle);
    
    var patron = overdueNoticeItems[i][1];
    if (i + 1 != overdueNoticeItems.length) {
      var nextPatron = overdueNoticeItems[i+1][1];
      if (patron == nextPatron) {
        newPatron = false;
      } else {
        var summary = body.appendParagraph('Please return your material as soon as possible. You will be billed the replacement cost for any material not returned within 5 weeks of the due date, and your account will be blocked. The Library uses a collection agency for any outstanding balances over $75.00.');
        summary.setAttributes(branchstyle);
        var closing = body.appendParagraph('For more information, visit SCLSNJ.org to view your account online. If you have any questions, you can call the ' +
                                           contact + ', or email ask@sclibnj.org.');
        closing.setAttributes(branchstyle);
        body.appendPageBreak();
        newPatron = true;
      }
    } else {
      var summary = body.appendParagraph('Please return your material as soon as possible. You will be billed the replacement cost for any material not returned within 5 weeks of the due date, and your account will be blocked. The Library uses a collection agency for any outstanding balances over $75.00.');
      summary.setAttributes(branchstyle);
      var closing = body.appendParagraph('For more information, visit SCLSNJ.org to view your account online. If you have any questions, you can call the ' +
                                         contact + ', or email ask@sclibnj.org.');
      closing.setAttributes(branchstyle);
      body.appendPageBreak();
    }
  }
  var updateDoc = DriveApp.getFileById(docId);
  updateDoc.getParents().next().removeFile(updateDoc);
  DriveApp.getFolderById('1D_FguqscVCeSFOLH4CWLzr60DUmtk7jG').addFile(updateDoc);
  var docURL = 'https://drive.google.com/open?id=' + updateDoc.getId();
  GmailApp.sendEmail('lhoffman@sclibnj.org, clarkson@sclibnj.org, mketselm@sclibnj.org', 'Today\'s Notices', docURL); 
}


function titleCase(str) {
  var words =  str.toLowerCase().split(' ');
  Logger.log(words);
  var update = '';
  for (var w = 0; w < words.length; w++) {
    var word = words[w];
    var re = /(^[a-z])/;
    var newstr = word.replace(re, function($0) {
      var letter = $0;
      return letter.toUpperCase();
    });
    if (w == 0) {
      update = newstr;
    } else {
      update += ' ' + newstr;
    }
  }
  return update;
}

