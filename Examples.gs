/*
 * The following are example functions which utilize the GmailUtils and
 * DriveUtils libraries to convert Gmail messages into PDF files.
 */


// Iterate through all starred messages with the label 'Expenses' (up to 5
// messages per call), convert each to a PDF, and save them to Google Drive,
// then unstar each message after archival.

function saveExpenses() {
  GmailUtils.processStarred('label:Expenses', 5, function(message) {

    // create a pdf of the message
    var pdf = GmailUtils.messageToPdf(message);

    // prefix the pdf filename with a date string
    pdf.setName(GmailUtils.formatDate(message, 'yyyyMMdd - ') + pdf.getName());

    // save the converted file to the 'Expenses' folder within Google Drive
    DriveUtils.getFolder('Expenses').createFile(pdf);

    // signal that we are done with this email and it will be marked as read
    return true;
  });
}


// Convert the 10 most recent unread messages in the "promotions" category
// into a single pdf, mark them as read, and email the digest to yourself.

function unreadPromotionsDigest() {
  var to = Session.getActiveUser().getEmail(),
      subject = 'Latest promotional emails',
      body = "here's your unread promotional emails:",
      messages = [];

  // limit this to 10 threads (may be more than 10 messages)
  GmailUtils.processUnread('category:promotions', 10, function(message) {
    messages.push(message);
    return true;
  });

  // convert the group of messages into a single pdf
  var pdf = GmailUtils.messageToPdf(messages, {filename: 'recent_promos.pdf'});

  // email the converted document
  GmailApp.sendEmail(to, subject, body, {attachments: [pdf]});
}
