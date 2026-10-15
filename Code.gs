/**
 * Participants Details for Certificate — Google Sheets backend
 * Spreadsheet: 1UEoJeWsLZk0iJsXXXy4ey6yoHwCboRH4YD7mDL3mDBo
 *
 * Deploy as a Web app:
 * Execute as: Me
 * Who has access: Anyone (or the access setting required by your organization)
 *
 * The frontend should POST JSON as text/plain to the deployed /exec URL.
 */

const SPREADSHEET_ID = '1UEoJeWsLZk0iJsXXXy4ey6yoHwCboRH4YD7mDL3mDBo';
const SHEET_NAME = 'Participant Records';
const PDF_FOLDER_ID = '1cNSHlQ3-ZkjYWm9mfwaV9-salfP5bM7w';

const HEADERS = [
  'Record ID',
  'Saved At',
  'Salutation',
  'Full Name (Hindi)',
  'Full Name (English)',
  'Record Type',
  'Workshop Title (Hindi)',
  'Workshop Title (English)',
  'Workshop Start Date',
  'Workshop End Date',
  'Designation (Hindi)',
  'Designation (English)',
  'School / Institution (Hindi)',
  'School / Institution (English)',
  'KVS Region (Hindi)',
  'KVS Region (English)',
  'Gender',
  'Category',
  'CPD Hours',
  'Certificate Date',
  'TNA / Reflection / Feedback Status',
  'Remarks (Hindi)',
  'Remarks (English)',
  'PDF File URL',
  'PDF Status'
];

function doGet() {
  return jsonOutput({
    ok: true,
    service: 'Participants Details for Certificate',
    message: 'Backend is deployed and reachable. Send a POST request to save a participant record.'
  });
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (payload.action === 'ping') {
      return jsonOutput({ ok: true, message: 'Connection successful.' });
    }
    if (payload.action !== 'save') {
      return jsonOutput({ ok: false, error: 'Unsupported action. Use action: save.' });
    }

    const record = payload.record || {};
    const missing = requiredFields().filter(function (key) {
      const value = record[key];
      return value === undefined || value === null ||
        (Array.isArray(value) ? value.length === 0 : String(value).trim() === '');
    });
    if (missing.length) {
      return jsonOutput({ ok: false, error: 'Missing required fields.', fields: missing });
    }
    if (String(record.endDate) < String(record.startDate)) {
      return jsonOutput({ ok: false, error: 'End date cannot be earlier than start date.' });
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
    ensureHeaders(sheet);

    const recordId = String(record.id || Utilities.getUuid());
    const savedAt = new Date();
    const row = [
      recordId,
      savedAt,
      joinValue(record.salutation),
      record.fullNameHi,
      record.fullNameEn,
      record.recordType,
      record.workshopTitleHi,
      record.workshopTitleEn,
      record.startDate,
      record.endDate,
      record.designationHi,
      record.designationEn,
      record.schoolHi,
      record.schoolEn,
      record.regionHi,
      record.regionEn,
      record.gender,
      joinValue(record.category),
      record.cpd,
      record.certDate,
      record.attendance,
      record.notesHi,
      record.notesEn,
      '',
      ''
    ];

    // If the same ID is submitted again, update the existing row rather than duplicate it.
    const idValues = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues()
      : [];
    let targetRow = -1;
    for (let i = 0; i < idValues.length; i++) {
      if (String(idValues[i][0]) === recordId) {
        targetRow = i + 2;
        break;
      }
    }
    if (targetRow > 0) {
      row[1] = sheet.getRange(targetRow, 2).getValue() || savedAt;
    } else {
      targetRow = sheet.getLastRow() + 1;
    }

    // Create the PDF in the requested Drive folder. Keep the record in Sheets even if PDF creation fails.
    let pdfUrl = '';
    let pdfStatus = '';
    try {
      const pdfFile = createParticipantPdf_(record, recordId);
      pdfUrl = pdfFile.getUrl();
      pdfStatus = 'PDF created';
    } catch (pdfErr) {
      pdfStatus = 'PDF failed: ' + String(pdfErr && pdfErr.message ? pdfErr.message : pdfErr);
    }
    row[row.length - 2] = pdfUrl;
    row[row.length - 1] = pdfStatus;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

    return jsonOutput({
      ok: true,
      message: 'Record saved to Google Sheets. ' + pdfStatus,
      recordId: recordId,
      row: targetRow,
      pdfUrl: pdfUrl,
      pdfStatus: pdfStatus
    });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}


function createParticipantPdf_(record, recordId) {
  const folder = DriveApp.getFolderById(PDF_FOLDER_ID);
  const safeName = String(record.fullNameEn || record.fullNameHi || 'Participant')
    .replace(/[\\\\/:*?"<>|#%{}~&]/g, '-').replace(/\\s+/g, ' ').trim();
  const fileName = 'Certificate_Details_' + safeName + '_' + recordId + '.pdf';

  // Replace a prior PDF for the same record ID so edits do not leave duplicate PDFs.
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const labels = [
    ['Record ID', recordId],
    ['Salutation / संबोधन', joinValue(record.salutation)],
    ['Full Name (Hindi) / पूरा नाम', record.fullNameHi],
    ['Full Name (English)', record.fullNameEn],
    ['Record Type / श्रेणी', record.recordType],
    ['Workshop Title (Hindi) / कार्यशाला शीर्षक', record.workshopTitleHi],
    ['Workshop Title (English)', record.workshopTitleEn],
    ['Workshop Start Date / प्रारंभ तिथि', record.startDate],
    ['Workshop End Date / समाप्ति तिथि', record.endDate],
    ['Designation (Hindi) / पदनाम', record.designationHi],
    ['Designation (English)', record.designationEn],
    ['School / Institution (Hindi) / विद्यालय', record.schoolHi],
    ['School / Institution (English)', record.schoolEn],
    ['KVS Region (Hindi) / संभाग', record.regionHi],
    ['KVS Region (English)', record.regionEn],
    ['Gender', record.gender],
    ['Category / श्रेणी', joinValue(record.category)],
    ['CPD Hours', record.cpd],
    ['Certificate Date / प्रमाण-पत्र तिथि', record.certDate],
    ['TNA / Reflection / Feedback Status', record.attendance],
    ['Remarks (Hindi) / टिप्पणी', record.notesHi],
    ['Remarks (English)', record.notesEn]
  ];
  const rows = labels.map(function(pair) {
    return '<tr><th>' + escapeHtml_(pair[0]) + '</th><td>' + escapeHtml_(pair[1]) + '</td></tr>';
  }).join('');
  const html = '<!doctype html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;color:#202433;font-size:10pt;margin:28px}' +
    'h1{color:#650909;font-size:19pt;margin-bottom:4px}p{color:#555;font-size:10pt}' +
    'table{border-collapse:collapse;width:100%;margin-top:18px}th,td{border:1px solid #d8dbe5;padding:8px;text-align:left;vertical-align:top;word-break:break-word}' +
    'th{width:36%;background:#f7eeee;color:#650909} .footer{margin-top:18px;font-size:8pt;color:#777}' +
    '</style></head><body><h1>Kendriya Vidyalaya Sangathan — ZIET Mumbai</h1>' +
    '<p>Participants Details for Certificate / प्रमाण-पत्र हेतु प्रतिभागी विवरण</p>' +
    '<table>' + rows + '</table><div class="footer">Generated automatically from the submitted participant record.</div></body></html>';
  const pdfBlob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF).setName(fileName);
  return folder.createFile(pdfBlob);
}

function escapeHtml_(value) {
  return String(value == null ? '' : value).replace(/[&<>"]/g, function(ch) {
    return {'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;'}[ch];
  });
}

function requiredFields() {
  return [
    'salutation', 'fullNameHi', 'fullNameEn', 'recordType',
    'workshopTitleHi', 'workshopTitleEn', 'startDate', 'endDate',
    'designationHi', 'designationEn', 'schoolHi', 'schoolEn',
    'regionHi', 'regionEn', 'gender', 'category', 'cpd',
    'certDate', 'attendance', 'notesHi', 'notesEn'
  ];
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#650909')
      .setFontColor('#ffffff');
    sheet.autoResizeColumns(1, HEADERS.length);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length))
    .getDisplayValues()[0];
  if (existing[0] !== HEADERS[0]) {
    throw new Error('The first row of "' + SHEET_NAME + '" does not match the expected headers.');
  }
  // Add any newly introduced columns without disturbing existing records.
  HEADERS.forEach(function(header, i) {
    if (String(existing[i] || '').trim() !== header) {
      sheet.getRange(1, i + 1).setValue(header)
        .setFontWeight('bold').setBackground('#650909').setFontColor('#ffffff');
    }
  });
  sheet.setFrozenRows(1);
}

function joinValue(value) {
  return Array.isArray(value) ? value.join(', ') : String(value == null ? '' : value);
}

function jsonOutput(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
