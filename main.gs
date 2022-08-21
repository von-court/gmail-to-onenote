const cfg = {
    'log': false,
    'onLabel': 'OneNote',
    'onenoteMail': 'me@onenote.com',
    'fieldGetterMap': { bcc: 'Bcc', cc: 'Cc', date: 'Date', from: 'From', replyto: 'ReplyTo', subject: 'Subject', to: 'To' },
    'hdrFields': 'From, To, Cc, Date',
    'hdrCss': 'border-bottom:1px solid #ccc;padding-bottom:1em;margin-bottom:1em;',
    'nacct': 1,
    'footerTag': 'Sent with Gmail To OneNote',
    'sentLabel': '',
    'keepSent': false
  };
const sentLabel = cfg.sentLabel ? GmailApp.getUserLabelByName(cfg.sentLabel) : undefined;
const userMail = Session.getActiveUser().getEmail();
let logSheet;

/**
*  Kick it!
*/
function execute() {
  try {
    // Catch errors, usually connecting to Gmail
    let labels = GmailApp.getUserLabels();
    labels.
      filter(label => label.getName().startsWith(cfg.onLabel)).
      map(label => { 
        label.getThreads().
          map((thread) => transferLastMessage_(thread, label));
      });
  } catch (e) {
    Logger.log('%s: %s %s', e.name, e.message, e.stack);
  }
}

/**
*  Transfer the last message of a thread
*/
function transferLastMessage_(thread, label) {
  let subject = thread.getFirstMessageSubject();
  let lastMsg = thread.getMessages().pop();
  let alreadySent = (lastMsg.getTo() == cfg.onenoteMail)
  if (!alreadySent) {

    let msgLabels = thread.getLabels().map((x) => x.getName());
    subject += msgLabels.map(createTargetNotebookTag_).join('');

    const hdrFields = cfg.hdrFields.toLowerCase().split(/[, ]+/);
    header = composeHeader_(hdrFields, lastMsg);

    /* Disable link creation to prevent OneNote to automatically append a URL 
      preview image of the Gmail login page

    header += createLink_(mailUrl_(lastMsg.getId(), cfg.nacct));
    */

    let body = lastMsg.getBody();
    header && (body = `<div style="${cfg.hdrCss}">${header.replace(/\n/g, "<br/>")}</div>${body}`);
    body = truncateByBytesUTF8(body, 199 * 1024);
    body = body + '\n[' + cfg.footerTag + ']';

    let attachments = lastMsg.getAttachments();

    /* Alternative to make Gmail send emojis as well (https://stackoverflow.com/questions/66077675/how-to-use-emoji-in-email-subject-using-google-apps-script)

    GmailApp.sendEmail(
      cfg.onenoteMail,
      `=?UTF-8?B?${Utilities.base64Encode(Utilities.newBlob(subject).getBytes())}?=`,
      '',
      {
        htmlBody: body,
        attachments: attachments
      }
    );
    */

    MailApp.sendEmail(
      cfg.onenoteMail,
      subject,
      '',
      {
        htmlBody: body,
        attachments: attachments
      }
    );

    log_(subject);
    sentLabel && threads.addLabel(sentLabel);
    cfg.keepSent || deleteSentMessage_();
  }
  thread.removeLabel(label);
  Utilities.sleep(1000);
}

/**
*  Delete just sent message
*/
const deleteSentMessage_ = () => {
  Utilities.sleep(3000);
  let sentResults = GmailApp.search("label:sent " + cfg.footerTag, 0, 1);
  if (sentResults) {
    let sentMsg = sentResults[0].getMessages().pop();
    (sentMsg.getTo() == cfg.onenoteMail) && sentMsg.moveToTrash();
  }
}

/**
 * Compose the header
 */
const composeHeader_ = (fields, lastMsg) => {
  let header = fields.
    reduce((header, field) => appendField_(header, field, lastMsg), '');
  return htmlEncode_(header);
}

/**
*  Append field to header
*/
const appendField_ = (header, field, lastMsg) => {
  let fldName = cfg.fieldGetterMap[field];
  if (fldName) {
    let fldValue = eval(`lastMsg.get${fldName}()`);
    fldValue && (header += `${fldName}: ${fldValue}\n`);
  }
  return header;
}

/**
*  Identify target notebook and return appropriate tag
*/
function createTargetNotebookTag_(label) {
  var msgLabelPath = label.split('/');
  if (msgLabelPath[0] == cfg.onLabel && msgLabelPath[1]) {
    return ' @' + msgLabelPath[1];
  }
  return '';
}

/**
*  Add an entry to the logsheet
*/
function log_(subject) {
  if (cfg.log) {
    try {
      if (!logSheet) {
        logSheetParts = cfg.log.split(':');
        var ss = SpreadsheetApp.openById(logSheetParts[0]);
        if (logSheetParts.length > 1) {
          logSheet = ss.getSheetByName(logSheetParts[1]);
        } else {
          logSheet = ss.getSheets()[0];
        }
      }
      logSheet.appendRow([new Date(),
      'OneNote:' + userMail,
        subject]);
    } catch (e) {
      Logger.log('%s at line %s: %s', e.name, e.lineNumber, e.message);
    }
  } else {
    Logger.log(subject);
  }
}

/**
*  Create a new log sheet if it does not exist
*/
function createLogSheet_() {
  let ss = SpreadsheetApp.create("Gmail to OneNote Log");
  let log_id = ss.getId();
  Logger.log(`Log sheet created with ID: ${log_id}\nURL: ${ss.getUrl()}`);
  let sheet = ss.getSheets()[0];
  sheet.appendRow(['Date', 'Source', 'Message']);
  return log;
}

/**
*  Create a HTML link
*/
function createLink_(url, text) {
  if (text == undefined) {
    text = url;
  }
  return sprintf_('<a href="%s">%s</a>', url, text);
}

/**
*  Create a Gmail url
*/
function mailUrl_(msgId, user) {
  return sprintf_('https://mail.google.com/mail/u/%s/#inbox/%s', user, msgId);
}

/**
*  Simple alternative for php sprintf-like text replacements
*  Each '%s' in the format is replaced by an additional parameter
*  E.g. sprintf_( '<a href="%s">%s</a>', url, text ) results in '<a href="url">text</a>'
*/
function sprintf_(format) {
  for (var i = 1; i < arguments.length; i++) {
    format = format.replace(/%s/, arguments[i]);
  }
  return format;
}

/**
*  Encode special HTML characters
*  From: http://jsperf.com/htmlencoderegex
*/
function htmlEncode_(html) {
  return html.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
*  Convert characters to bytes
*  From: https://stackoverflow.com/a/1516420/5630865
*/
function toBytesUTF8(chars) {
  return unescape(encodeURIComponent(chars));
}

/**
*  Convert bytes to characters
*  From: https://stackoverflow.com/a/1516420/5630865
*/
function fromBytesUTF8(bytes) {
  return decodeURIComponent(escape(bytes));
}

/**
*  Truncate characters to specific byte count
*  From: https://stackoverflow.com/a/1516420/5630865
*/
function truncateByBytesUTF8(chars, n) {
  var bytes = toBytesUTF8(chars).substring(0, n);
  while (true) {
    try {
      return fromBytesUTF8(bytes);
    } catch (e) { };
    bytes = bytes.substring(0, bytes.length - 1);
  }
}

/**
 * Remove non-UTF8 characters
 */
function cleanString(input) {
    let output = "";
    // adapted to german character set
    output = input.replace(/[^a-zA-Z0-9ÄÖÜäöüß_@]/g, ' ');
    return output;
}
