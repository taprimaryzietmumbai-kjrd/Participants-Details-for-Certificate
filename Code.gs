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
  'Remarks (English)'
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
      record.notesEn
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
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
      targetRow = sheet.getLastRow();
    }

    return jsonOutput({
      ok: true,
      message: targetRow > 0 ? 'Record saved to Google Sheets.' : 'Record saved.',
      recordId: recordId,
      row: targetRow
    });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err && err.message ? err.message : err) });
  }
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
    throw new Error('The first row of "' + SHEET_NAME + '" does not match the expected headers. Rename or clear that sheet before first use.');
  }
}

function joinValue(value) {
  return Array.isArray(value) ? value.join(', ') : String(value == null ? '' : value);
}

function jsonOutput(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
