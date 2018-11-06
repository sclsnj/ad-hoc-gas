function getEmailLogs() {
  // look for a Drive file that has the right name; if it's not there, create it
  var now = new Date();
  // look in email for yesterday's data reports
  var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);  
  var dateRegExp = new RegExp(yesterday);
  var label = GmailApp.getUserLabelByName("xCARL Email Logs");
  var threads = label.getThreads();
  var text = '';
  var noticeFlag = false;
  var bounceFlag = false;
  for (var i = 0; i < threads.length; i++) {
    var message = threads[i].getMessages()[0];
    var messageDate = message.getDate();
    if (messageDate < now && messageDate > yesterday) {
      var subject = message.getSubject();
      if (subject.match(/notice/)) {
        var body = message.getBody();
        if (body.match(/Sent (\d{1,3}) out of /)) {
          var notices = body.match(/Sent (\d{1,3}) out of /);
          Logger.log(notices);
          Logger.log(notices[1]);
          noticeFlag = true;
        }
      } else if (subject.match(/bounce/)) {
        var body = message.getBody();
        if (body.match(/(\d{1,3}) Messages found/)) {
          var bounce = body.match(/(\d{1,3}) Messages found/);
          Logger.log(bounce[1]);
          bounceFlag = true;
        }
      }
      if (bounceFlag && noticeFlag) {
        break;
      }
    }
  }
  if (!bounceFlag) {
    text += 'No bounce log received. \n';
  } else {
    if (bounce[1] > 35) {
      text += 'High bounce: ' + bounce[1] + '\n';
    } else if (bounce[1] < 5) {
      text += 'Low bounce: ' + bounce[1] + '\n';
    }
  }
  if (!noticeFlag) {
    text += 'No notice log received. \n';
  } else {
    if (notices[1] > 300) {
      text += 'High notices: ' + notices[1] + '\n';
    } else if (notices[1] < 50) {
      text += 'Low notices: ' + notices[1] + '\n';
    }
  }
  if (text) {
    GmailApp.sendEmail('lhoffman@sclibnj.org', 'Email log issue', text);
  }
}
