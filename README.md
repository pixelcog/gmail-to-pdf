Gmail To PDF
============

A helpful collection of Google Apps Script utilities which allow you to
painlessly transform Gmail messages into PDFs.
[View on Google Apps Script](https://script.google.com/d/1qdkT9ShXl4VWO9XvKefcxmH_oRJe31MPDyIDsOKyGidKr-GHBpULLtvx/edit?usp=sharing)

This handles all of the heavy-lifting needed to preserve all attachments,
inline-images, remote images, backgrounds, etc. Also included are a few helpful
methods for creating virtual queues within your Gmail labels, using either
unread status or starred status as a signifier.


## Example Usage

Convert all unread items within the label "Expenses" into PDFs and save them to
Google Drive:

```javascript
// iterate through all unread messages with the label 'Expenses' (limited to 5 at a time)
GmailUtils.processUnread('label:Expenses', 5, function(msg) {

  // convert the message to a pdf
  pdf = GmailUtils.messageToPdf(msg);

  // save the converted file to the 'Expenses' folder within Google Drive
  DriveUtils.getFolder('Expenses').createFile(pdf);

  // signal that we are done with this email and it will be marked as read
  return true;
}
```

Convert all starred emails to PDFs and mail them to _handle@example.com_ before
unstarring them.

```javascript
// iterate through all unread messages with the label 'Expenses'
GmailUtils.processStarred(function(msg) {

  // convert the message to a pdf without including any attachments
  pdf = GmailUtils.messageToPdf(msg, {name: 'email.pdf', embedAttachments: false});

  // email the converted document to handle@example.com
  var to = "handle@example.com",
      subject = 'Converted: ' + msg.getSubject(),
      body = "here's your converted pdf:";

  GmailApp.sendEmail(to, subject, body, {attachments: [pdf]});

  // return true so the message will be unstarred
  return true;
});
```


LICENSE
=======

The MIT License (MIT)

Copyright (c) 2015 PixelCog Inc.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.