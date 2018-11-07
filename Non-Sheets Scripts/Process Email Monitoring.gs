/* This function is scheduled to run each morning to look for CARL email notice and bounce monitoring emails, 
 * and to flag them when they don't run correctly or when the numbers are lower or higher than expected.
 * This is dependent on monitoring emails being received in Gmail to start with, and has to run as a script in
 * Google Drive for the user who is receiving the emails in their Gmail account.
 *
 * To prep, we have our monitoring emails set up to be filtered on receipt into a general xCARL Email Logs label.
 * That allows for more focused processing of just emails that have to do with CARL, rather than looking through
 * the entire inbox.
 */

function getEmailLogs() {
  // Look in Gmail for emails that have been labeled as in the email log messages.
  var now = new Date();
  var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);  
  var dateRegExp = new RegExp(yesterday);
  var label = GmailApp.getUserLabelByName("xCARL Email Logs");
  var threads = label.getThreads(0,4);
  var highVolume, lowVolume, noticeFailure, highBounce, lowBounce, bounceIncomplete;
  // For the first email message in each message thread (there's only one per thread anyway), see if it's a
  // notice or a bounce email.
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var message = thread.getMessages()[0];
    var messageDate = message.getDate();
    if (messageDate < now && messageDate > yesterday) {
      var subject = message.getSubject();
      // Evaluate the body text of notice emails for volume anomalies, other failure issues.
      if (subject.match(/notice/)) {
        var body = message.getBody();
        var notices = body.match(/Sent (\d{1,3}) out of (\d{1,3})/);
        if (notices.length > 0) {
          if (notices[1] !== notices[2]) {
            noticeFailure = true;
          } else {
            if (notices[1] > 300) {
              highVolume = true;
            } else if (notices[1] < 50) {
              lowVolume = true;
            }
          }
        } else {
          noticeFailure = true;
        }
        var smtp = body.match('SMTP Server: smtp-relay.gmail.com');
        if (smtp.length = 0) {
          noticeFailure = true;
        }
      // Evaluate the body text of bounce emails for volume anomalies, other failure issues.
      } else if (subject.match(/bounce/)) {
        var body = message.getBody();
        if (body.match(/Entering *(ErisOutputManager )Shutdown/i) && body.match(/Done Parsing Email/i)) {
        } else {
          bounceIncomplete = true;
        }
        var bounce = body.match(/(\d{1,3}) Messages found/);
        if (bounce.length > 0) {
          if (bounce[1] > 35) {
            highBounce = true;
          } else if (bounce[1] < 5) {
            lowBounce = true;
          }
        } else {
          bounceIncomplete = true;
        }
      }
    }
    // Based on the flags set above, add a new label to the email thread under consideration, push the 
    // thread back into the inbox, and mark unread to make sure it's visible.
    if (highBounce) {
      var newlabel = GmailApp.getUserLabelByName('xHigh Bounce');
      thread.addLabel(newlabel).moveToInbox().markUnread();
      highBounce = false;
    }
    if (lowBounce) {
      var newlabel = GmailApp.getUserLabelByName('xLow Bounce');
      thread.addLabel(newlabel).moveToInbox().markUnread();
      lowBounce = false;
    }
    if (bounceIncomplete) {
      var newlabel = GmailApp.getUserLabelByName('xBounce Incomplete');
      thread.addLabel(newlabel).moveToInbox().markUnread();
      bounceIncomplete = false;
    }
    if (highVolume) {
      var newlabel = GmailApp.getUserLabelByName('xHigh Volume');
      thread.addLabel(newlabel).moveToInbox().markUnread();
      highVolume = false;
    }
    if (lowVolume) {
      var newlabel = GmailApp.getUserLabelByName('xLow Volume');
      threads.addLabel(newlabel).moveToInbox().markUnread();
      lowVolume = false;
    }
    if (noticeFailure) {
      var newlabel = GmailApp.getUserLabelByName('xNotice Failure');
      thread.addLabel(newlabel).moveToInbox().markUnread();
      noticeFailure = false;
    }
  }    
}
