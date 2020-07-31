/*
 * GmailUtils
 * ==========
 *
 * A collection of utilities for use with Google Apps Email and the GmailApp object, primarily
 * focused on archiving messages as PDFs.
 *
 * To utilize this library, select Resources > Libraries... and enter the following project key:
 * MsE3tErxE9G0z6EMfGmUGqVVaKzeOjMwH
 */

/**
 * Iterate through all messages matching the given query.
 *
 * @method eachMessage
 * @param {string} query (optional, default 'in:inbox')
 * @param {number} limit (optional, default 10)
 * @param {function} callback
 */
function eachMessage(query, limit, callback) {
  if (typeof query == 'function') {
    callback = query;
    query = null;
    limit = null;
  }
  if (typeof limit == 'function') {
    callback = limit;
    limit = null;
  }
  if (typeof callback != 'function') {
    throw "No callback provided";
  }
  limit = parseInt(limit) || 10;
  query = query || 'in:inbox';

  var threads = GmailApp.search(query, 0, limit);
  for (var t=0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m=0; m < messages.length; m++) {
      callback(messages[m]);
    }
  }
}

/**
 * Iterate through all starred messages which match the given query. When the callback returns a
 * positive value, the message is unstarred.
 *
 * @method processStarred
 * @param {string} query (optional, default 'is:starred')
 * @param {number} limit (optional, default 10)
 * @param {function} callback
 */
function processStarred(query, limit, callback) {
  if (typeof query == 'function') {
    callback = query;
    query = null;
    limit = null;
  }
  if (typeof limit == 'function') {
    callback = limit;
    limit = null;
  }
  if (typeof callback != 'function') {
    throw "No callback provided";
  }
  query = (query ? query + ' AND ' : '') + 'is:starred';

  eachMessage(query, limit, function(message) {
    message.isStarred() && !message.isInTrash() && callback(message) && message.unstar();
  });
}

/**
 * Iterate through all unread messages which match the given query. When the callback returns a
 * positive value, the message is marked as read.
 *
 * @method processUnread
 * @param {string} query (optional, default 'is:unread')
 * @param {number} limit (optional, default 10)
 * @param {function} callback
 */
function processUnread(query, limit, callback) {
  if (typeof query == 'function') {
    callback = query;
    query = null;
    limit = null;
  }
  if (typeof limit == 'function') {
    callback = limit;
    limit = null;
  }
  if (typeof callback != 'function') {
    throw "No callback provided";
  }
  query = (query ? query + ' AND ' : '') + 'is:unread';

  eachMessage(query, limit, function(message) {
    message.isUnread() && !message.isInTrash() && callback(message) && message.markRead();
  });
}

/**
 * Wrapper for Utilities.formatDate() which provides sensiblea defaults
 *
 * @method formatDate
 * @param {string} date //Change from 'message' to 'date' to make it not limited to gmail messages, then can be invoked anywhere in the project.
 * @param {string} format
 * @param {string} timezone
 * @return {string} Formatted date
 */
function formatDate(date, format, timezone) {
  timezone = timezone || localTimezone_();
  format = format || "MMMMM dd, yyyy 'at' h:mm a '" + timezone + "'";
  return Utilities.formatDate(new Date(date), timezone, format)
}

/**
 * Determine whether a message has a pdf attached to it and if so, return it
 *
 * @method messageGetPdfAttachment
 * @param {GmailMessage} message GmailMessage object
 * @return {Blob|boolean} Blob on success, else false
 */
function messageGetPdfAttachment(message) {
  var attachments = message.getAttachments();
  for (var i=0; i < attachments.length; i++) {
    if (attachments[i].getContentType() == 'application/octet-stream') {
      attachments[i].setContentTypeFromExtension();
    }
    if (attachments[i].getContentType() == 'application/pdf') {
      return attachments[i].copyBlob();
    }
  }
  return false;
}

/**
 * Convert a Gmail message or thread to a PDF and return it as a blob
 *
 * @method messageToPdf
 * @param {GmailMessage|GmailThread} messages GmailMessage or GmailThread object (or an array of such objects)
 * @return {Blob}
 */
function messageToPdf(messages, opts) {
  return messageToHtml(messages, opts).getAs('application/pdf');
}

/**
 * Convert a Gmail message or thread to a HTML and return it as a blob
 *
 * @method messageToHtml
 * @param {GmailMessage|GmailThread} messages GmailMessage or GmailThread object (or an array of such objects)
 * @param {Object} options
 * @return {Blob}
 */
function messageToHtml(messages, opts) {
  opts = opts || {};
  defaults_(opts, {
    includeHeader: true,
    includeAttachments: true,
    embedAttachments: true,
    embedRemoteImages: true,
    embedInlineImages: true,
    embedAvatar: true,
    width: 700,
    filename: null
  });

  if (!(messages instanceof Array)) {
    messages = isa_(messages, 'GmailThread') ? messages.getMessages() : [messages];
  }
  if (!messages.every(function(obj){ return isa_(obj, 'GmailMessage'); })) {
    throw "Argument must be of type GmailMessage or GmailThread.";
  }
  var name = opts.filename || sanitizeFilename_(messages[messages.length-1].getSubject()) + '.html';
  var html = '<html>\n' +
             '<style type="text/css">\n' +
             'body{padding:0 10px;min-width:' + opts.width + 'px;-webkit-print-color-adjust: exact;}' +
             'body>dl.email-meta{font-family:"Helvetica Neue",Helvetica,Arial,sans-serif;font-size:14px;padding:0 0 10px;margin:0 0 5px;border-bottom:1px solid #ddd;page-break-before:always}' +
             'body>dl.email-meta:first-child{page-break-before:auto}' +
             'body>dl.email-meta dt{color:#808080;float:left;width:60px;clear:left;text-align:right;overflow:hidden;text-overf‌low:ellipsis;white-space:nowrap;font-style:normal;font-weight:700;line-height:1.4}' +
             'body>dl.email-meta dd{margin-left:70px;line-height:1.4}' +
             'body>dl.email-meta dd a{color:#808080;font-size:0.85em;text-decoration:none;font-weight:normal}' +
             'body>dl.email-meta dd.avatar{float:right}' +
             'body>dl.email-meta dd.avatar img{max-height:72px;max-width:72px;border-radius:36px}' +
             'body>dl.email-meta dd.strong{font-weight:bold}' +
             'body>div.email-attachments{font-size:0.85em;color:#999}' +
             '</style>\n' +
             '<body>\n';

  for (var m=0; m < messages.length; m++) {
    var message = messages[m],
        subject = message.getSubject(),
        avatar = null,
        date = formatDate(message.getDate()),
        from = formatEmails_(message.getFrom()),
        to   = formatEmails_(message.getTo()),
        cc   = formatEmails_(message.getCc()),
        bcc  = formatEmails_(message.getBcc()),
        body = message.getBody();

    if (opts.includeHeader) {
      if (opts.embedAvatar && (avatar = emailGetAvatar(from))) {
        avatar = '<dd class="avatar"><img src="' + renderDataUri_(avatar) + '" /></dd> ';
      } else {
        avatar = '';
      }
      html += '<dl class="email-meta">\n' +
              '<dt>From:</dt>' + avatar + ' <dd class="strong">' + from + '</dd>\n' +
              '<dt>Subject:</dt> <dd>' + subject + '</dd>\n' +
              '<dt>Date:</dt> <dd>' + date + '</dd>\n' +
              '<dt>To:</dt> <dd>' + to + '</dd>\n';
    }
    // Appending cc and bcc if they exist
    if(isRealValue(cc)){
          html += '<dt>cc:</dt> <dd>' + cc + '</dd>\n';
    }
    if(isRealValue(bcc)){
          html += '<dt>bcc:</dt> <dd>' + bcc + '</dd>\n';
    }
    html += '</dl>\n';
    if (opts.embedRemoteImages) {
      body = embedHtmlImages_(body);
    }
    if (opts.embedInlineImages) {
      body = embedInlineImages_(body, message.getRawContent());
    }
    if (opts.includeAttachments) {
      var attachments = message.getAttachments();
      if (attachments.length > 0) {
        body += '<br />\n<strong>Attachments:</strong>\n' +
                '<div class="email-attachments">\n';

        for (var a=0; a < attachments.length; a++) {
          var filename = attachments[a].getName();
          var imageData;

          if (opts.embedAttachments && (imageData = renderDataUri_(attachments[a]))) {
            body += '<img src="' + imageData + '" alt="&lt;' + filename + '&gt;" /><br />\n';
          } else {
            body += '&lt;' + filename + '&gt;<br />\n';
          }
        }
        body += '</div>\n';
      }
    }
    html += body;
  }
  html += '</body>\n</html>';

  return Utilities.newBlob(html, 'text/html', name);
}

/**
 * Returns the name associated with an email string, or the domain name of the email.
 *
 * @method emailGetName
 * @param {string} email
 * @return {string} name or domain name
 */
function emailGetName(email) {
  return email.replace(/^<?(?:[^<\(]+@)?([^<\(,]+?|)(?:\s?[\(<>,].*|)$/i, '$1') || 'Unknown';
}

/**
 * Attempt to download an image representative of the email address provided. Using gravatar or
 * apple touch icons as appropriate.
 *
 * @method emailGetAvatar
 * @param {string} email
 * @return {Blob|boolean} Blob object or false
 */
function emailGetAvatar(email) {
  re = /[a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?/gi
  if (!(email = email.match(re)) || !(email = email[0].toLowerCase())) {
    return false;
  }
  var domain = email.split('@')[1];
  var avatar = fetchRemoteFile_('http://www.gravatar.com/avatar/' + md5_(email) + '?s=128&d=404');
  if (!avatar && ['gmail','hotmail','yahoo.'].every(function(s){ return domain.indexOf(s) == -1 })) {
    avatar = fetchRemoteFile_('http://' + domain + '/apple-touch-icon.png') ||
             fetchRemoteFile_('http://' + domain + '/apple-touch-icon-precomposed.png');
  }
  return avatar;
}

/**
 * Download and embed all images referenced within an html document as data uris
 *
 * @param {string} html
 * @return {string} Html with embedded images
 */
function embedHtmlImages_(html) {
  // process all img tags
  html = processImageTags(html);
  // process all style attributes
  html = processStyleAttributes(html);
  // process all style tags
  html = processStyleTags(html);
  return html;
}

/**
 * Download and embed all img tags
 *
 * @param {string} html
 * @return {string} Html with embedded images
 */
function processImageTags(html){
    return html.replace(/(<img[^>]+src=)(["'])((?:(?!\2)[^\\]|\\.)*)\2/gi, function(m, tag, q, src) {
        // Logger.log('Processing image src: ' + src);
        return tag + q + (renderDataUri_(src) || src) + q;
    });
}

/**
 * Download and embed all HTML Style Attributes
 *
 * @param {string} html
 * @return {string} Html with embedded style attributes
 */
function processStyleAttributes(html){
    return html.replace(/(<[^>]+style=)(["'])((?:(?!\2)[^\\]|\\.)*)\2/gi, function(m, tag, q, style) {
        style = style.replace(/url\((\\?["']?)([^\)]*)\1\)/gi, function(m, q, url) {
            return 'url(' + q + (renderDataUri_(url) || url) + q + ')';
        });
        return tag + q + style + q;
    });
}

/**
 * Download and embed all HTML Style Tags
 *
 * @param {string} html
 * @return {string} Html with embedded style tags
 */
function processStyleTags(html){
    return html.replace(/(<style[^>]*>)(.*?)(?:<\/style>)/gi, function(m, tag, style, end) {
    style = style.replace(/url\((["']?)([^\)]*)\1\)/gi, function(m, q, url) {
      return 'url(' + q + (renderDataUri_(url) || url) + q + ')';
    });
    return tag + style + end;
  });
}

/**
 * Extract and embed all inline images (experimental)
 *
 * @param {string} html Message body
 * @param {string} raw Unformatted message contents
 * @return {string} Html with embedded images
 */
function embedInlineImages_(html, raw) {
  var images = [];

  // locate all inline content ids
  raw.replace(/<img[^>]+src=(?:3D)?(["'])cid:((?:(?!\1)[^\\]|\\.)*)\1/gi, function(m, q, cid) {
    images.push(cid);
    return m;
  });

  // extract all inline images
  images = images.map(function(cid) {
    var cidIndex = raw.search(new RegExp("Content-ID ?:.*?" + cid, 'i'));
    if (cidIndex === -1) return null;

    var prevBoundaryIndex = raw.lastIndexOf("\r\n--", cidIndex);
    var nextBoundaryIndex = raw.indexOf("\r\n--", prevBoundaryIndex+1);
    var part = raw.substring(prevBoundaryIndex, nextBoundaryIndex);

    var encodingLine = part.match(/Content-Transfer-Encoding:.*?\r\n/i)[0];
    var encoding = encodingLine.split(":")[1].trim();
    if (encoding != "base64") return null;

    var contentTypeLine = part.match(/Content-Type:.*?\r\n/i)[0];
    var contentType = contentTypeLine.split(":")[1].split(";")[0].trim();

    var startOfBlob = part.indexOf("\r\n\r\n");
    var blobText = part.substring(startOfBlob).replace("\r\n","");

    return Utilities.newBlob(Utilities.base64Decode(blobText), contentType, cid);
  }).filter(function(i){return i});

  // process all img tags which reference "attachments"
  return processImgAttachments(html);
}

/**
 * Download and embed all HTML Inline Image Attachments
 *
 * @param {string} html
 * @return {string} Html with inline image attachments
 */
function processImgAttachments(html){
    return html.replace(/(<img[^>]+src=)(["'])(\?view=att(?:(?!\2)[^\\]|\\.)*)\2/gi, function(m, tag, q, src) {
        return tag + q + (renderDataUri_(images.shift()) || src) + q;
   });
}

/**
 * Convert an image into a base64-encoded data uri.
 *
 * @param {Blob|string} Blob object containing an image file or a remote url string
 * @return {string} Data uri
 */
function renderDataUri_(image) {
  if (typeof image == 'string' && !(isValidUrl_(image) && (image = fetchRemoteFile_(image)))) {
    return null;
  }
  if (isa_(image, 'Blob') || isa_(image, 'GmailAttachment')) {
    if (image.getContentType() != null) {
      var type = image.getContentType().toLowerCase();
      var data = Utilities.base64Encode(image.getBytes());
      if (type.indexOf('image') == 0) {
        return 'data:' + type + ';base64,' + data;
      }
    } 
  }
  return null;
}

/**
 * Fetch a remote file and return as a Blob object on success
 *
 * @param {string} url
 * @return {Blob}
 */
function fetchRemoteFile_(url) {
  try {
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    return response.getResponseCode() == 200 ? response.getBlob() : null;
  } catch (e) {
    return null;
  }
}

/**
 * Validate a url string (taken from jQuery)
 *
 * @param {string} url
 * @return {boolean}
 */
function isValidUrl_(url) {
  return /^(https?|ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(\#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i.test(url);
}

/**
 * Sanitize a filename by filtering out characters not allowed in most filesystems
 *
 * @param {string} filename
 * @return {string}
 */
function sanitizeFilename_(filename) {
  return filename.replace(/[\/\?<>\\:\*\|":\x00-\x1f\x80-\x9f]/g, '');
}

/**
 * Turn emails of the form "<handle@domain.tld>" into 'mailto:' links.
 *
 * @param {string} emails
 * @return {string}
 */
function formatEmails_(emails) {
  var pattern = new RegExp(/<(((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)>/i);
  return emails.replace(pattern, function(match, handle) {
    return '<a href="mailto:' + handle + '">' + handle + '</a>';
  });
}

/**
 * Test class name for Google Apps Script objects. They have no constructors so we must test them
 * with toString.
 *
 * @param {Object} obj
 * @param {string} class
 * @return {boolean}
 */
function isa_(obj, class) {
  return typeof obj == 'object' && typeof obj.constructor == 'undefined' && obj.toString() == class;
}

/**
 * Assign default attributes to an object.
 *
 * @param {Object} options
 * @param {Object} defaults
 */
function defaults_(options, defaults) {
  for (attr in defaults) {
    if (!options.hasOwnProperty(attr)) {
      options[attr] = defaults[attr];
    }
  }
}

/**
 * Get our current timezone string (or GMT if it cannot be determined)
 *
 * @return {string}
 */
function localTimezone_() {
  var timezone = CalendarApp.getDefaultCalendar().getTimeZone();
  return timezone.length ? timezone : 'GMT';
}

/**
* Check if value is not null or undefined
*
* @param {Object} obj
* @return {boolean} true if object is not null or undefined
*/
function isRealValue(obj) {
 return obj && obj !== 'null' && obj !== 'undefined';
}

/**
 * Create an MD5 hash of a string and return the reult as hexadecimal.
 *
 * @param {string} str
 * @return {string}
 */
function md5_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str).reduce(function(str,chr) {
    chr = (chr < 0 ? chr + 256 : chr).toString(16);
    return str + (chr.length==1?'0':'') + chr;
  },'');
}
