/**
 * A collection of utilities for use with Google Apps Email and the GmailApp object, primarily
 * focused on archiving messages as PDFs.
 *
 * @module GmailUtils
 */

var GmailUtils = new function() {

  /**
   * Iterate through all messages matching the given query.
   *
   * @method eachMessage
   * @param {String} query (optional, default 'in:inbox')
   * @param {Integer} limit (optional, default 10)
   * @param {Function} callback
   */
  this.eachMessage = function(query, limit, callback) {
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
  };

  /**
   * Iterate through all starred messages which match the given query. When the callback returns a
   * positive value, the message is unstarred.
   *
   * @method processStarred
   * @param {String} query (optional, default 'is:starred')
   * @param {Integer} limit (optional, default 10)
   * @param {Function} callback
   */
  this.processStarred = function(query, limit, callback) {
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

    this.eachMessage(query, limit, function(message) {
      message.isStarred() && !message.isInTrash() && callback(message) && message.unstar();
    });
  };

  /**
   * Iterate through all unread messages which match the given query. When the callback returns a
   * positive value, the message is marked as read.
   *
   * @method processUnread
   * @param {String} query (optional, default 'is:unread')
   * @param {Integer} limit (optional, default 10)
   * @param {Function} callback
   */
  this.processUnread = function(query, limit, callback) {
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

    this.eachMessage(query, limit, function(message) {
      message.isUnread() && !message.isInTrash() && callback(message) && message.markRead();
    });
  };

  /**
   * Wrapper for Utilities.formatDate() which provides sensible defaults
   *
   * @method formatDate
   * @param {String} message
   * @param {String} format
   * @param {String} timezone
   * @return {String} Formatted date
   */
  this.formatDate = function(message, format, timezone) {
    timezone = timezone || localTimezone();
    format = format || "MMMMM dd, yyyy 'at' h:mm a '" + timezone + "'";
    return Utilities.formatDate(message.getDate(), timezone, format)
  };

  /**
   * Determine whether a message has a pdf attached to it and if so, return it
   *
   * @method messageGetPdfAttachment
   * @param {Object} message GmailMessage object
   * @return {Object|Boolean} Blob on success, else false
   */
  this.messageGetPdfAttachment = function(message) {
    var attachments = message.getAttachments();
    for (var i=0; i < attachments.length; i++) {
      if (attachments[i].getContentType() == 'application/pdf') {
        return attachments[i].copyBlob();
      }
    }
    return false;
  };

  /**
   * Convert a Gmail message or thread to a PDF and return it as a blob
   *
   * @method messageToPdf
   * @param {Object} messages GmailMessage or GmailThread object (or an array of such objects)
   * @return {Object} Blob
   */
  this.messageToPdf = function(messages, opts) {
    return this.messageToHtml(messages, opts).getAs('application/pdf');
  };

  /**
   * Convert a Gmail message or thread to a HTML and return it as a blob
   *
   * @method messageToHtml
   * @param {Object} messages GmailMessage or GmailThread object (or an array of such objects)
   * @param {Object} options Self explanitory, see source
   * @return {Object} Blob
   */
  this.messageToHtml = function(messages, opts) {
    opts = opts || {};
    defaults(opts, {
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
      messages = isa(messages, 'GmailThread') ? messages.getMessages() : [messages];
    }
    if (!messages.every(function(obj){ return isa(obj, 'GmailMessage'); })) {
      throw "Argument must be of type GmailMessage or GmailThread.";
    }
    var name = opts.filename || sanitizeFilename(messages[messages.length-1].getSubject()) + '.html';
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
          date = this.formatDate(message),
          from = formatEmails(message.getFrom()),
          to   = formatEmails(message.getTo()),
          body = message.getBody();

      if (opts.includeHeader) {
        if (opts.embedAvatar && (avatar = this.emailGetAvatar(from))) {
          avatar = '<dd class="avatar"><img src="' + renderDataUri(avatar) + '" /></dd> ';
        } else {
          avatar = '';
        }
        html += '<dl class="email-meta">\n' +
                '<dt>From:</dt>' + avatar + ' <dd class="strong">' + from + '</dd>\n' +
                '<dt>Subject:</dt> <dd>' + subject + '</dd>\n' +
                '<dt>Date:</dt> <dd>' + date + '</dd>\n' +
                '<dt>To:</dt> <dd>' + to + '</dd>\n' +
                '</dl>\n';
      }
      if (opts.embedRemoteImages) {
        body = embedHtmlImages(body);
      }
      if (opts.embedInlineImages) {
        body = embedInlineImages(body, message.getRawContent());
      }
      if (opts.includeAttachments) {
        var attachments = message.getAttachments();
        if (attachments.length > 0) {
          body += '<br />\n<strong>Attachments:</strong>\n' +
                  '<div class="email-attachments">\n';

          for (var a=0; a < attachments.length; a++) {
            var filename = attachments[a].getName();
            var imageData;

            if (opts.embedAttachments && (imageData = renderDataUri(attachments[a]))) {
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
  };

  /**
   * Returns the name associated with an email string, or the domain name of the email.
   *
   * @method emailGetName
   * @param {String} email
   * @return {String} name or domain name
   */
  this.emailGetName = function(email) {
    return email.replace(/^<?(?:[^<\(]+@)?([^<\(,]+?|)(?:\s?[\(<>,].*|)$/i, '$1') || 'Unknown';
  };

  /**
   * Attempt to download an image representative of the email address provided. Using gravatar or
   * apple touch icons as appropriate.
   *
   * @method emailGetAvatar
   * @param {String} email
   * @return {Object|Boolean} Blob object or false
   */
  this.emailGetAvatar = function(email) {
    re = /[a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?/gi
    if (!(email = email.match(re)) || !(email = email[0].toLowerCase())) {
      return false;
    }
    var domain = email.split('@')[1];
    var avatar = fetchRemoteFile('http://www.gravatar.com/avatar/' + md5(email) + '?s=128&d=404');
    if (!avatar && ['gmail','hotmail','yahoo.'].every(function(s){ return domain.indexOf(s) == -1 })) {
      avatar = fetchRemoteFile('http://' + domain + '/apple-touch-icon.png') ||
               fetchRemoteFile('http://' + domain + '/apple-touch-icon-precomposed.png');
    }
    return avatar;
  };

  /**
   * Download and embed all images referenced within an html document as data uris
   *
   * @param {String} html
   * @return {String} Html with embedded images
   */
  function embedHtmlImages(html) {
    // process all img tags
    html = html.replace(/(<img[^>]+src=)(["'])((?:(?!\2)[^\\]|\\.)*)\2/gi, function(m, tag, q, src) {
      // Logger.log('Processing image src: ' + src);
      return tag + q + (renderDataUri(src) || src) + q;
    });
    // process all style attributes
    html = html.replace(/(<[^>]+style=)(["'])((?:(?!\2)[^\\]|\\.)*)\2/gi, function(m, tag, q, style) {
      style = style.replace(/url\((\\?["']?)([^\)]*)\1\)/gi, function(m, q, url) {
        return 'url(' + q + (renderDataUri(url) || url) + q + ')';
      });
      return tag + q + style + q;
    });
    // process all style tags
    html = html.replace(/(<style[^>]*>)(.*?)(?:<\/style>)/gi, function(m, tag, style, end) {
      style = style.replace(/url\((["']?)([^\)]*)\1\)/gi, function(m, q, url) {
        return 'url(' + q + (renderDataUri(url) || url) + q + ')';
      });
      return tag + style + end;
    });
    return html;
  }

  /**
   * Extract and embed all inline images (experimental)
   *
   * @param {String} html Message body
   * @param {String} raw Unformatted message contents
   * @return {String} Html with embedded images
   */
  function embedInlineImages(html, raw) {
    var images = [];

    // locate all inline content ids
    raw.replace(/<img[^>]+src=(["'])cid:((?:(?!\1)[^\\]|\\.)*)\1/gi, function(m, q, cid) {
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
    return html.replace(/(<img[^>]+src=)(["'])(\?view=att(?:(?!\2)[^\\]|\\.)*)\2/gi, function(m, tag, q, src) {
      return tag + q + (renderDataUri(images.shift()) || src) + q;
    });
  }

  /**
   * Convert an image into a base64-encoded data uri.
   *
   * @param {Object|String} Blob object containing an image file or a remote url string
   * @return {String} Data uri
   */
  function renderDataUri(image) {
    if (typeof image == 'string' && !(isValidUrl(image) && (image = fetchRemoteFile(image)))) {
      return null;
    }
    if (isa(image, 'Blob') || isa(image, 'GmailAttachment')) {
      var type = image.getContentType().toLowerCase();
      var data = Utilities.base64Encode(image.getBytes());
      if (type.indexOf('image') == 0) {
        return 'data:' + type + ';base64,' + data;
      }
    }
    return null;
  }

  /**
   * Fetch a remote file and return as a Blob object on success
   *
   * @param {String} url
   * @return {Object} Blob
   */
  function fetchRemoteFile(url) {
    // Logger.log('Fetching url: ' + url);
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true})
    return response.getResponseCode() == 200 ? response.getBlob() : null;
  }

  /**
   * Validate a url string (taken from jQuery)
   *
   * @param {String} url
   * @return {Boolean}
   */
  function isValidUrl(url) {
    return /^(https?|ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(\#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i.test(url);
  }

  /**
   * Sanitize a filename by filtering out characters not allowed in most filesystems
   *
   * @param {String} filename
   * @return {String}
   */
  function sanitizeFilename(filename) {
    return filename.replace(/[\/\?<>\\:\*\|":\x00-\x1f\x80-\x9f]/g, '');
  }

  /**
   * Turn emails of the form "<handle@domain.tld>" into 'mailto:' links.
   *
   * @param {String} emails
   * @return {String}
   */
  function formatEmails(emails) {
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
   * @param {String} class
   * @return {Boolean}
   */
  function isa(obj, class) {
    return typeof obj == 'object' && typeof obj.constructor == 'undefined' && obj.toString() == class;
  }

  /**
   * Assign default attributes to an object.
   *
   * @param {Object} options
   * @param {Object} defaults
   */
  function defaults(options, defaults) {
    for (attr in defaults) {
      if (!options.hasOwnProperty(attr)) {
        options[attr] = defaults[attr];
      }
    }
  }

  /**
   * Get our current timezone string (or GMT if it cannot be determined)
   *
   * @return {String}
   */
  function localTimezone() {
    var timezone = new Date().toTimeString().match(/\(([a-z0-9]+)\)/i);
    return timezone.length ? timezone[1] : 'GMT';
  }

  /**
   * Create an MD5 hash of a string and return the reult as hexadecimal.
   *
   * @param {String} str
   * @return {String}
   */
  function md5(str) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str).reduce(function(str,chr) {
      chr = (chr < 0 ? chr + 256 : chr).toString(16);
      return str + (chr.length==1?'0':'') + chr;
    },'');
  }
};

