// backup all flagged messages as specified within the user properties

function archiveGmailMessages() {
  var props = PropertiesService.getUserProperties(),
      query = props.getProperty('query'),
      limit = props.getProperty('limit'),
      sendTo = props.getProperty('send_to'),
      saveTo = props.getProperty('save_to');

  GmailUtils.processStarred(query, limit, function(msg) {
    Logger.log('Processing Email: ' + msg.getSubject());

    var date = GmailUtils.formatDate(msg, 'yyyyMMdd'),
        from = GmailUtils.emailGetName(msg.getFrom()),
        file = GmailUtils.messageGetPdfAttachment(msg) || GmailUtils.messageToPdf(msg),
        name = [date, from, file.getName()].join(' - '),
        body = 'From: ' + msg.getFrom() + '<br />\n' +
               'Date: ' + GmailUtils.formatDate(msg);

    file.setName(name);

    var saved = saveTo && !!DriveUtils.getFolder(saveTo).createFile(file);
    var sent  = sendTo && !!MailApp.sendEmail({
      to: sendTo,
      body: body,
      subject: msg.getSubject(),
      attachments: [file],
      noReply: true
    });

    return saved || sent;
  });
}
