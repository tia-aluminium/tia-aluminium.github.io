const SHEET_NAME = 'RFQ Submissions';
const NOTIFY_EMAIL = 'rsaini65@gmail.com';
const MAX_FILE_BYTES = 25 * 1024 * 1024; // Gmail's attachment cap

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.company_website) {
      // Honeypot field — real users never fill this in. Pretend success
      // so bots don't notice and adapt, but skip the sheet/email entirely.
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = getOrCreateSheet();
    let fileBlob = null;
    let fileName = '';

    if (data.file && data.file.data) {
      const bytes = Utilities.base64Decode(data.file.data);
      if (bytes.length > MAX_FILE_BYTES) {
        // Reject before touching Sheet/Mail at all — an oversized file
        // would never transport via Gmail's attachment cap anyway.
        return ContentService.createTextOutput(JSON.stringify({
          status: 'error',
          message: 'File exceeds the 25 MB limit.',
        })).setMimeType(ContentService.MimeType.JSON);
      }
      fileBlob = Utilities.newBlob(bytes, data.file.mimeType, data.file.name);
      fileName = data.file.name;
    }

    sheet.appendRow([
      new Date(),
      data.fullName || '',
      data.email || '',
      data.phone ? "'" + data.phone : '',
      data.quantity || '',
      data.sector || '',
      data.details || '',
      fileName,
    ]);

    const bodyLines = [
      'New RFQ submitted on the Tia Aluminium Profiles website:',
      '',
      'Name: ' + (data.fullName || '-'),
      'Email: ' + (data.email || '-'),
      'Phone: ' + (data.phone || '-'),
      'Estimated Quantity: ' + (data.quantity || '-'),
      'Sector: ' + (data.sector || '-'),
      'Details: ' + (data.details || '-'),
      'Blueprint: ' + (fileName ? 'Attached (' + fileName + ')' : 'No file attached'),
    ];

    const mailOptions = {};
    if (fileBlob) mailOptions.attachments = [fileBlob];

    MailApp.sendEmail(NOTIFY_EMAIL, 'New RFQ — Tia Aluminium Profiles', bodyLines.join('\n'), mailOptions);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Full Name', 'Email', 'Phone', 'Quantity', 'Sector', 'Details', 'Blueprint Filename']);
  }
  return sheet;
}
