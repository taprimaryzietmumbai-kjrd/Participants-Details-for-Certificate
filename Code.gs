
/**
 * Participants Details for Certificate — ZIET Mumbai
 * Saves participant details, photo, signature and PDF.
 *
 * Spreadsheet: 1UEoJeWsLZk0iJsXXXy4ey6yoHwCboRH4YD7mDL3mDBo
 * Sheet: Participant Records
 * Drive folder for PDFs, photos and signatures:
 * 1cNSHlQ3-ZkjYWm9mfwaV9-salfP5bM7w
 *
 * Web app deployment:
 * Execute as: Me
 * Access: Anyone (or the access level required by your organisation)
 */

const SPREADSHEET_ID = '1UEoJeWsLZk0iJsXXXy4ey6yoHwCboRH4YD7mDL3mDBo';
const SHEET_NAME = 'Participant Records';
const FILES_FOLDER_ID = '1cNSHlQ3-ZkjYWm9mfwaV9-salfP5bM7w';
const MASTER_SPREADSHEET_ID = '1Uuqoz2KFpTcw9d2HI02fLmFPopTaLnKWwZ9aR-uYCeY';

const HEADERS = [
  'S.No.',
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
  'PDF Status',
  'Employee Code',
  'Year of Joining',
  'Participant WhatsApp',
  'Participant Email',
  'Vidyalaya Email',
  'Participant Photo URL',
  'Participant Signature URL',
  'Email Status',
  'Certificate Receipt Status'
];

/* -------------------- WEB APP -------------------- */

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      service: 'Participants Details for Certificate',
      message: 'Backend is reachable.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = (e && e.parameter && e.parameter.payload) ||
      (e && e.postData && e.postData.contents) || '{}';
    const payload = JSON.parse(raw);

    if (payload.action === 'ping') {
      return responseOutput_({ ok: true, message: 'Connection successful.' });
    }

    if (payload.action === 'receiptAcknowledged') {
      return responseOutput_(recordCertificateReceipt_(payload));
    }

    if (payload.action !== 'save') {
      return responseOutput_({
        ok: false,
        error: 'Unsupported action. Use action: save.'
      });
    }

    const record = payload.record || {};
    const missing = requiredFields_().filter(function(key) {
      const value = record[key];
      return value === undefined || value === null ||
        (Array.isArray(value) ? value.length === 0 : String(value).trim() === '');
    });

    if (missing.length) {
      return responseOutput_({
        ok: false,
        error: 'Missing required fields.',
        fields: missing
      });
    }

    if (String(record.endDate) < String(record.startDate)) {
      return responseOutput_({
        ok: false,
        error: 'End date cannot be earlier than start date.'
      });
    }

    // Photo and signature are mandatory in this form.
    if (!record.participantPhotoDataUrl || !record.signatureDataUrl) {
      return responseOutput_({
        ok: false,
        error: 'Participant photo and signature are required.'
      });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    ensureHeaders_(sheet);
    const columns = getHeaderMap_(sheet);
    const recordId = String(record.id || Utilities.getUuid());
    const now = new Date();

    // Find the existing record by its Record ID column, not column A.
    const idCol = columns['record id'];
    let targetRow = 0;
    let savedAt = now;
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getDisplayValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === recordId) {
          targetRow = i + 2;
          savedAt = sheet.getRange(targetRow, columns['saved at']).getValue() || now;
          break;
        }
      }
    }

    if (!targetRow) targetRow = lastRow + 1;

    // Preserve any existing values in the row when updating a record.
    const rowWidth = sheet.getLastColumn();
    let row = targetRow <= lastRow
      ? sheet.getRange(targetRow, 1, 1, rowWidth).getValues()[0]
      : new Array(rowWidth).fill('');

    function put(header, value) {
      const col = columns[normalise_(header)];
      if (col) row[col - 1] = value;
    }

    // Serial number is only assigned for a new record.
    if (targetRow > lastRow) {
      put('S.No.', Math.max(0, lastRow - 1) + 1);
    }

    put('Record ID', recordId);
    put('Saved At', savedAt);
    put('Salutation', joinValue_(record.salutation));
    put('Full Name (Hindi)', record.fullNameHi);
    put('Full Name (English)', record.fullNameEn);
    put('Record Type', record.recordType);
    put('Workshop Title (Hindi)', record.workshopTitleHi);
    put('Workshop Title (English)', record.workshopTitleEn);
    put('Workshop Start Date', record.startDate);
    put('Workshop End Date', record.endDate);
    put('Designation (Hindi)', record.designationHi);
    put('Designation (English)', record.designationEn);
    put('School / Institution (Hindi)', record.schoolHi);
    put('School / Institution (English)', record.schoolEn);
    put('KVS Region (Hindi)', record.regionHi);
    put('KVS Region (English)', record.regionEn);
    put('Gender', record.gender);
    put('Category', joinValue_(record.category));
    put('CPD Hours', record.cpd);
    put('Certificate Date', record.certDate);
    put('TNA / Reflection / Feedback Status', record.attendance);
    put('Remarks (Hindi)', record.notesHi);
    put('Remarks (English)', record.notesEn);
    put('Employee Code', record.employeeCode);
    put('Year of Joining', record.yearOfJoining);
    put('Participant WhatsApp', record.participantWhatsapp);
    put('Participant Email', record.participantEmail);
    put('Vidyalaya Email', record.vidyalayaEmail);

    // Save the row first, so an image/PDF error cannot lose the participant details.
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();

    // SAVE PHOTO AND SIGNATURE TO GOOGLE DRIVE.
    // The functions return the Drive URLs, which are written to the sheet.
    let photoUrl = '';
    let signatureUrl = '';

    try {
      photoUrl = saveImageDataUrl_(
        record.participantPhotoDataUrl,
        'Participant_Photo_' + safeFilePart_(recordId),
        recordId,
        'photo'
      );
      put('Participant Photo URL', photoUrl);

      signatureUrl = saveImageDataUrl_(
        record.signatureDataUrl,
        'Participant_Signature_' + safeFilePart_(recordId),
        recordId,
        'signature'
      );
      put('Participant Signature URL', signatureUrl);

      // Save the image URLs immediately.
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
      SpreadsheetApp.flush();

    } catch (imageError) {
      put('PDF Status', 'Image saving failed: ' + errorMessage_(imageError));
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
      SpreadsheetApp.flush();

      return responseOutput_({
        ok: false,
        error: 'Participant details were saved, but photo/signature saving failed: ' +
          errorMessage_(imageError),
        recordId: recordId,
        row: targetRow
      });
    }

    // Optional employee-code write-back to the master certificate spreadsheet.
    let masterSync = { ok: false, message: 'Not attempted.' };
    try {
      masterSync = syncEmployeeCodeToMaster_(record);
    } catch (syncError) {
      masterSync = { ok: false, message: errorMessage_(syncError) };
    }

    // Create a PDF after the image files have been saved.
    let pdfUrl = '';
    let pdfStatus = '';
    try {
      const pdf = createParticipantPdf_(
        record, recordId, photoUrl, signatureUrl
      );
      pdfUrl = pdf.getUrl();
      pdfStatus = 'PDF created';
    } catch (pdfError) {
      pdfStatus = 'PDF failed: ' + errorMessage_(pdfError);
    }

    put('PDF File URL', pdfUrl);
    put('PDF Status', pdfStatus);
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();

    return responseOutput_({
      ok: true,
      message: 'Participant record, photo and signature saved. ' + pdfStatus,
      recordId: recordId,
      row: targetRow,
      photoUrl: photoUrl,
      signatureUrl: signatureUrl,
      pdfUrl: pdfUrl,
      pdfStatus: pdfStatus,
      masterSync: masterSync
    });

  } catch (err) {
    return responseOutput_({
      ok: false,
      error: errorMessage_(err)
    });
  }
}

/* -------------------- DRIVE IMAGE SAVING -------------------- */

/**
 * Accepts a data URL such as data:image/jpeg;base64,...
 * Saves it in the configured Drive folder and returns its URL.
 * Replaces an existing same-name file to avoid duplicates on updates.
 */
function saveImageDataUrl_(dataUrl, baseName, recordId, kind) {
  if (!dataUrl || typeof dataUrl !== 'string') {
    throw new Error('No image data received for ' + kind + '.');
  }

  const parsed = parseImageDataUrl_(dataUrl);
  const folder = DriveApp.getFolderById(FILES_FOLDER_ID);
  const fileName = baseName + '.' + parsed.extension;

  // Remove only a prior image with this exact name.
  const oldFiles = folder.getFilesByName(fileName);
  while (oldFiles.hasNext()) {
    oldFiles.next().setTrashed(true);
  }

  const blob = Utilities.newBlob(
    parsed.bytes,
    parsed.mimeType,
    fileName
  );

  const file = folder.createFile(blob);
  return file.getUrl();
}

function parseImageDataUrl_(dataUrl) {
  const match = dataUrl.match(
    /^data:(image\/(?:jpeg|jpg|png|webp));base64,([A-Za-z0-9+/=\r\n]+)$/
  );

  if (!match) {
    throw new Error(
      'Invalid image data URL. Supported formats are JPEG, PNG and WebP.'
    );
  }

  const mimeType = match[1].toLowerCase();
  const base64 = match[2].replace(/\s/g, '');
  const bytes = Utilities.base64Decode(base64);

  if (!bytes || !bytes.length) {
    throw new Error('The received ' + mimeType + ' image is empty.');
  }

  // Prevent unexpectedly large image payloads.
  if (bytes.length > 8 * 1024 * 1024) {
    throw new Error('Image is larger than 8 MB. Please capture it again.');
  }

  let extension = 'jpg';
  if (mimeType === 'image/png') extension = 'png';
  if (mimeType === 'image/webp') extension = 'webp';

  return { bytes: bytes, mimeType: mimeType, extension: extension };
}

/* -------------------- SPREADSHEET HEADER MAPPING -------------------- */

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    styleHeaders_(sheet, HEADERS.length);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const existingKeys = existing.map(normalise_);

  // Never replace or reorder existing columns.
  // Append any missing required headers at the end.
  HEADERS.forEach(function(header) {
    if (existingKeys.indexOf(normalise_(header)) === -1) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(header);
      existingKeys.push(normalise_(header));
    }
  });

  sheet.setFrozenRows(1);
  styleHeaders_(sheet, sheet.getLastColumn());
}

function styleHeaders_(sheet, count) {
  sheet.getRange(1, 1, 1, count)
    .setFontWeight('bold')
    .setBackground('#650909')
    .setFontColor('#ffffff');
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(
    1, 1, 1, sheet.getLastColumn()
  ).getDisplayValues()[0];

  const map = {};
  headers.forEach(function(header, index) {
    const key = normalise_(header);
    if (key) map[key] = index + 1;
  });

  // Support common older header variations.
  const aliases = {
    'participant whatsapp': [
      'participant whatsapp no.',
      'participant whatsapp number',
      'participant whatsapp no'
    ],
    'participant photo url': [
      'photo url',
      'participant photo',
      'photo link'
    ],
    'participant signature url': [
      'signature url',
      'participant signature',
      'signature link'
    ]
  };

  Object.keys(aliases).forEach(function(canonical) {
    if (!map[canonical]) {
      aliases[canonical].some(function(alias) {
        if (map[normalise_(alias)]) {
          map[canonical] = map[normalise_(alias)];
          return true;
        }
        return false;
      });
    }
  });

  [
    'record id', 'saved at', 's.no.',
    'participant photo url', 'participant signature url',
    'pdf file url', 'pdf status'
  ].forEach(function(required) {
    if (!map[required]) {
      throw new Error('Required spreadsheet header not found: ' + required);
    }
  });

  return map;
}

function normalise_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

/* -------------------- PARTICIPANT PDF -------------------- */

function createParticipantPdf_(record, recordId, photoUrl, signatureUrl) {
  const folder = DriveApp.getFolderById(FILES_FOLDER_ID);
  const safeName = safeFilePart_(
    record.fullNameEn || record.fullNameHi || 'Participant'
  );
  const fileName = 'Certificate_Details_' + safeName + '_' + recordId + '.pdf';

  // Replace an older PDF for this same record.
  const oldPdfs = folder.getFilesByName(fileName);
  while (oldPdfs.hasNext()) oldPdfs.next().setTrashed(true);

  const labels = [
    ['Record ID', recordId],
    ['Salutation / संबोधन', joinValue_(record.salutation)],
    ['Full Name (Hindi) / पूरा नाम', record.fullNameHi],
    ['Full Name (English)', record.fullNameEn],
    ['Record Type / श्रेणी', record.recordType],
    ['Workshop Title (Hindi)', record.workshopTitleHi],
    ['Workshop Title (English)', record.workshopTitleEn],
    ['Workshop Start Date', record.startDate],
    ['Workshop End Date', record.endDate],
    ['Designation (Hindi)', record.designationHi],
    ['Designation (English)', record.designationEn],
    ['Employee Code', record.employeeCode],
    ['Year of Joining', record.yearOfJoining],
    ['Participant WhatsApp', record.participantWhatsapp],
    ['Participant Email', record.participantEmail],
    ['Vidyalaya Email', record.vidyalayaEmail],
    ['School / Institution (Hindi)', record.schoolHi],
    ['School / Institution (English)', record.schoolEn],
    ['KVS Region (Hindi)', record.regionHi],
    ['KVS Region (English)', record.regionEn],
    ['Gender', record.gender],
    ['Category', joinValue_(record.category)],
    ['CPD Hours', record.cpd],
    ['Certificate Date', record.certDate],
    ['TNA / Reflection / Feedback Status', record.attendance],
    ['Remarks (Hindi)', record.notesHi],
    ['Remarks (English)', record.notesEn],
    ['Participant Photo URL', photoUrl],
    ['Participant Signature URL', signatureUrl]
  ];

  const rows = labels.map(function(pair) {
    return '<tr><th>' + escapeHtml_(pair[0]) + '</th><td>' +
      escapeHtml_(pair[1]) + '</td></tr>';
  }).join('');

  const photoLink = photoUrl
    ? '<p><b>Participant photo:</b> <a href="' +
      escapeHtml_(photoUrl) + '">Open saved photo in Drive</a></p>'
    : '';
  const signatureLink = signatureUrl
    ? '<p><b>Participant signature:</b> <a href="' +
      escapeHtml_(signatureUrl) + '">Open saved signature in Drive</a></p>'
    : '';

  const html = '<!doctype html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,sans-serif;color:#202433;font-size:10pt;margin:28px}' +
    'h1{color:#650909;font-size:19pt;margin-bottom:4px}' +
    'p{color:#555;font-size:10pt;overflow-wrap:anywhere}' +
    'table{border-collapse:collapse;width:100%;margin-top:18px}' +
    'th,td{border:1px solid #d8dbe5;padding:8px;text-align:left;vertical-align:top;word-break:break-word}' +
    'th{width:36%;background:#f7eeee;color:#650909}' +
    '.footer{margin-top:18px;font-size:8pt;color:#777}' +
    '</style></head><body>' +
    '<h1>Kendriya Vidyalaya Sangathan — ZIET Mumbai</h1>' +
    '<p>Participants Details for Certificate / प्रमाण-पत्र हेतु प्रतिभागी विवरण</p>' +
    photoLink + signatureLink +
    '<table>' + rows + '</table>' +
    '<div class="footer">Generated automatically from the submitted participant record.</div>' +
    '</body></html>';

  const blob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .getAs(MimeType.PDF)
    .setName(fileName);

  return folder.createFile(blob);
}

/* -------------------- MASTER SHEET EMPLOYEE CODE SYNC -------------------- */

function syncEmployeeCodeToMaster_(record) {
  const employeeCode = String(record.employeeCode || '').trim();
  if (!employeeCode) {
    throw new Error('Employee Code is blank in the submitted record.');
  }

  const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const sheets = ss.getSheets();

  const norm = function(v) { return normalise_(v); };
  const nameKeys = [
    'Name of the Participant', 'Full Name (English)', 'Name', 'Participant Name'
  ];
  const schoolKeys = [
    'School', 'School / Institution (English)', 'Vidyalaya'
  ];
  const titleKeys = [
    'Title', 'Workshop Title (English)', 'Workshop Title', 'Course Title'
  ];
  const codeKeys = [
    'Code', 'Employee Code', 'Employee Code / कर्मचारी कोड'
  ];

  const matches = [];

  sheets.forEach(function(sheet) {
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) return;

    const headers = values[0].map(function(h) { return String(h || '').trim(); });
    const findCol = function(keys) {
      for (let i = 0; i < keys.length; i++) {
        const found = headers.findIndex(function(h) {
          return norm(h) === norm(keys[i]);
        });
        if (found >= 0) return found;
      }
      return -1;
    };

    const nameCol = findCol(nameKeys);
    const schoolCol = findCol(schoolKeys);
    const titleCol = findCol(titleKeys);
    const codeCol = findCol(codeKeys);
    if (nameCol < 0 || codeCol < 0) return;

    for (let r = 1; r < values.length; r++) {
      const nameOK = norm(values[r][nameCol]) === norm(record.fullNameEn);
      const schoolOK = schoolCol < 0 ||
        norm(values[r][schoolCol]) === norm(record.schoolHi) ||
        norm(values[r][schoolCol]) === norm(record.schoolEn);
      const titleOK = titleCol < 0 ||
        norm(values[r][titleCol]) === norm(record.workshopTitleEn);

      if (nameOK && schoolOK && titleOK) {
        matches.push({ sheet: sheet, row: r + 1, codeCol: codeCol });
      }
    }
  });

  if (matches.length !== 1) {
    throw new Error(matches.length
      ? 'More than one certificate row matched; employee code was not written.'
      : 'No unique participant row matched in the certificate master sheet.');
  }

  const match = matches[0];
  match.sheet.getRange(match.row, match.codeCol + 1)
    .setNumberFormat('@')
    .setValue(employeeCode);
  SpreadsheetApp.flush();

  const saved = String(
    match.sheet.getRange(match.row, match.codeCol + 1).getDisplayValue()
  ).trim();

  if (saved !== employeeCode) {
    throw new Error('Employee code write-back could not be verified.');
  }

  return {
    ok: true,
    sheet: match.sheet.getName(),
    row: match.row
  };
}

/* -------------------- CERTIFICATE RECEIPT ACKNOWLEDGEMENT -------------------- */

function recordCertificateReceipt_(payload) {
  const recordId = String(payload.recordId || '').trim();
  const employeeCode = String(payload.employeeCode || '').trim();
  const name = String(payload.participantName || '').trim();
  const title = String(payload.workshopTitle || '').trim();

  if (!recordId && !employeeCode && !name) {
    throw new Error('A participant identifier is required.');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Participant Records sheet not found.');

  ensureHeaders_(sheet);
  const columns = getHeaderMap_(sheet);
  const data = sheet.getDataRange().getDisplayValues();
  const matches = [];

  for (let r = 1; r < data.length; r++) {
    const idOK = recordId &&
      String(data[r][columns['record id'] - 1]).trim() === recordId;
    const codeOK = employeeCode && columns['employee code'] &&
      String(data[r][columns['employee code'] - 1]).trim() === employeeCode;
    const nameOK = name && columns['full name (english)'] &&
      String(data[r][columns['full name (english)'] - 1]).trim().toLowerCase() === name.toLowerCase();
    const titleOK = !title || !columns['workshop title (english)'] ||
      String(data[r][columns['workshop title (english)'] - 1]).trim().toLowerCase() === title.toLowerCase();

    if ((idOK || (!recordId && codeOK) ||
        (!recordId && !employeeCode && nameOK)) && titleOK) {
      matches.push(r + 1);
    }
  }

  if (matches.length !== 1) {
    throw new Error(matches.length
      ? 'Multiple participant rows match; receipt status was not changed.'
      : 'No unique participant row matched; receipt status was not changed.');
  }

  const statusCol = columns['certificate receipt status'];
  if (!statusCol) throw new Error('Certificate Receipt Status header not found.');

  sheet.getRange(matches[0], statusCol).setValue('Certificate received successfully');
  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Certificate receipt status saved.',
    row: matches[0]
  };
}

/* -------------------- COMMON HELPERS -------------------- */

function requiredFields_() {
  return [
    'salutation', 'fullNameHi', 'fullNameEn', 'recordType',
    'workshopTitleHi', 'workshopTitleEn', 'startDate', 'endDate',
    'designationHi', 'designationEn', 'employeeCode', 'yearOfJoining',
    'participantWhatsapp', 'participantEmail', 'vidyalayaEmail',
    'schoolHi', 'schoolEn', 'regionHi', 'regionEn', 'gender',
    'category', 'cpd', 'certDate', 'attendance', 'notesHi', 'notesEn'
  ];
}

function joinValue_(value) {
  return Array.isArray(value)
    ? value.join(', ')
    : String(value == null ? '' : value);
}

function safeFilePart_(value) {
  return String(value == null ? '' : value)
    .replace(/[\\/:*?"<>|#%{}~&]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 100);
}

function escapeHtml_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function(ch) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[ch];
  });
}

function errorMessage_(err) {
  return String(err && err.message ? err.message : err);
}

function responseOutput_(value) {
  const json = JSON.stringify(value).replace(/</g, '\\u003c');
  const html = '<!doctype html><html><head><meta charset="utf-8"></head>' +
    '<body><script>' +
    'window.top.postMessage({type:"participant-save-result",payload:' +
    json + '}, "*");' +
    '</script>Request processed.</body></html>';

  return HtmlService.createHtmlOutput(html);
}
