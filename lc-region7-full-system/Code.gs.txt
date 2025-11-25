/*****************************************************
 * สภาทนายความภาค 7 - Apps Script API สำหรับ Netlify
 * Backend: Google Apps Script
 * Frontend: Netlify (HTML/JS)
 * Database: Google Sheets (ไฟล์เดียว หลายชีต)
 * File Storage: Google Drive
 *
 * Module API:
 *  1) ประชาชน
 *      POST action=createCitizenCase
 *      GET  action=getCitizenStatus
 *
 *  2) ทนายความ
 *      POST action=createLawyerHelpCase
 *      GET  action=getLawyerHelpStatus
 *
 *  3) แอดมิน
 *      GET  action=adminSummary
 *      GET  action=adminListCases
 *      POST action=adminUpdateStatus
 *
 *****************************************************/

// ===== ตั้งค่าเชื่อมต่อ Google Workspace =====

// จากลิงก์ Google Sheet ของคุณ
const SPREADSHEET_ID = '1OjZWNt2gJ8hp1qmhrsoSbK0uUu651848Bu5Ng66r1Ps';

// จากลิงก์ Google Drive root folder ของคุณ
const DRIVE_ROOT_CITIZEN = '16cj5l3zK7rMT_sEAIsRZ6q1Pk8veUeQI';
const DRIVE_ROOT_LAWYER  = '16cj5l3zK7rMT_sEAIsRZ6q1Pk8veUeQI';

// ชื่อชีตหลัก (สร้างเพิ่มในไฟล์เก่าได้เลย)
const SHEET_CITIZEN = 'CITIZEN_CASES';
const SHEET_LAWYER  = 'LAWYER_CASES';
const SHEET_LOGS    = 'LOGS';

// รหัสแอดมินแบบง่าย (ต้องใช้ทั้งฝั่ง API และ admin.html ให้ตรงกัน)
const ADMIN_KEY = 'LC7-Admin-2025';   // แก้ให้เป็นรหัสที่คุณต้องการ

/**
 * โครงสร้างคอลัมน์ใน CITIZEN_CASES
 * สามารถสร้างชีตนี้ในไฟล์เก่า แล้วให้โค้ดเติมหัวตารางให้อัตโนมัติ
 */
const CITIZEN_HEADER = [
  'id',             // 0  รหัสภายใน (uuid)
  'reference_no',   // 1  เลขอ้างอิง C-...
  'created_at',     // 2  วันที่สร้าง
  'citizen_name',   // 3  ชื่อผู้ยื่น
  'phone',          // 4
  'line_id',        // 5
  'email',          // 6
  'case_type',      // 7
  'case_detail',    // 8
  'status',         // 9  RECEIVED / IN_PROGRESS / WAITING / CLOSED / REJECTED
  'drive_folder_id',// 10
  'last_update',    // 11
  'pdpa_consent',   // 12
  'admin_note'      // 13  หมายเหตุแอดมิน
];

/**
 * โครงสร้างคอลัมน์ใน LAWYER_CASES
 */
const LAWYER_HEADER = [
  'id',             // 0
  'reference_no',   // 1  เลขอ้างอิง L-...
  'created_at',     // 2
  'lawyer_name',    // 3
  'lawyer_license', // 4
  'region',         // 5
  'phone',          // 6
  'line_id',        // 7
  'email',          // 8
  'issue_type',     // 9
  'issue_detail',   // 10
  'status',         // 11
  'drive_folder_id',// 12
  'last_update',    // 13
  'pdpa_consent',   // 14
  'admin_note'      // 15
];

/**
 * โครงสร้างคอลัมน์ LOGS
 */
const LOG_HEADER = [
  'timestamp', // เวลา
  'user',      // PUBLIC / LAWYER / ADMIN / SYSTEM
  'action',    // CREATE_* / ADMIN_UPDATE_* / ERROR_*
  'module',    // CITIZEN / LAWYER / ADMIN
  'ref',       // reference_no
  'detail'     // รายละเอียด
];

// ========== ENTRY POINTS ==========

function doGet(e) {
  // กัน error เวลาเรียกจาก Editor โดยไม่มี e
  e = e || {};
  e.parameter = e.parameter || {};

  ensureSheetsStructure_();

  const action = (e.parameter.action || '').trim();

  // ping ทดสอบ API
  if (action === 'ping') {
    return jsonResponse({ ok: true, message: 'LC Region7 API is alive' });
  }

  // ===== ฝั่งประชาชน =====
  if (action === 'getCitizenStatus') {
    return handleGetCitizenStatus(e);
  }

  // ===== ฝั่งทนายความ =====
  if (action === 'getLawyerHelpStatus') {
    return handleGetLawyerHelpStatus(e);
  }

  // ===== ฝั่งแอดมิน: Summary Dashboard =====
  if (action === 'adminSummary') {
    if (!checkAdminKey_(e)) return jsonResponse({ error: 'unauthorized' });
    return handleAdminSummary(e);
  }

  // ===== ฝั่งแอดมิน: ดึงรายการเคส =====
  if (action === 'adminListCases') {
    if (!checkAdminKey_(e)) return jsonResponse({ error: 'unauthorized' });
    return handleAdminListCases(e);
  }

  return jsonResponse({ error: 'Unknown action (GET)' });
}

function doPost(e) {
  // กัน error เวลาเรียกจาก Editor โดยไม่มี e
  e = e || {};
  e.parameter = e.parameter || {};

  ensureSheetsStructure_();

  const action = (e.parameter.action || '').trim();

  // ===== ฝั่งประชาชน =====
  if (action === 'createCitizenCase') {
    return handleCreateCitizenCase(e);
  }

  // ===== ฝั่งทนายความ =====
  if (action === 'createLawyerHelpCase') {
    return handleCreateLawyerHelpCase(e);
  }

  // ===== ฝั่งแอดมิน: อัปเดตสถานะ =====
  if (action === 'adminUpdateStatus') {
    if (!checkAdminKey_(e)) return jsonResponse({ error: 'unauthorized' });
    return handleAdminUpdateStatus(e);
  }

  return jsonResponse({ error: 'Unknown action (POST)' });
}

// ========== HELPERS พื้นฐาน ==========

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/** ตรวจสอบ admin_key */
function checkAdminKey_(e) {
  const key = (e.parameter.admin_key || '').trim();
  return key && key === ADMIN_KEY;
}

/**
 * ตรวจและสร้างชีต + หัวตารางให้อัตโนมัติ
 */
function ensureSheetsStructure_() {
  const ss = getSpreadsheet_();

  // CITIZEN_CASES
  let citizenSheet = ss.getSheetByName(SHEET_CITIZEN);
  if (!citizenSheet) citizenSheet = ss.insertSheet(SHEET_CITIZEN);
  setupHeaderIfNeeded_(citizenSheet, CITIZEN_HEADER);

  // LAWYER_CASES
  let lawyerSheet = ss.getSheetByName(SHEET_LAWYER);
  if (!lawyerSheet) lawyerSheet = ss.insertSheet(SHEET_LAWYER);
  setupHeaderIfNeeded_(lawyerSheet, LAWYER_HEADER);

  // LOGS
  let logSheet = ss.getSheetByName(SHEET_LOGS);
  if (!logSheet) logSheet = ss.insertSheet(SHEET_LOGS);
  setupHeaderIfNeeded_(logSheet, LOG_HEADER);
}

/**
 * ถ้าแถวแรกว่าง → ใส่ header
 * ถ้ามีแล้วแต่ไม่ตรง → อัปเดตหัวตารางให้ตรงกับ config
 */
function setupHeaderIfNeeded_(sheet, headerArray) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headerArray);
    return;
  }
  const headRange   = sheet.getRange(1, 1, 1, headerArray.length);
  const currentHead = headRange.getValues()[0];
  const needFix = headerArray.some((h, i) => (currentHead[i] || '') !== h);
  if (needFix) {
    headRange.setValues([headerArray]);
  }
}

/** เขียน log ลงชีต LOGS */
function appendLog_(user, action, moduleName, ref, detail) {
  const ss       = getSpreadsheet_();
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  if (!logSheet) return;

  logSheet.appendRow([
    new Date(),
    user,
    action,
    moduleName,
    ref,
    detail
  ]);
}

/** แปลง status เป็นภาษาไทย (ใช้ร่วมกัน) */
function mapStatusToThai_(status) {
  switch (status) {
    case 'RECEIVED':    return 'รับเรื่องแล้ว';
    case 'IN_PROGRESS': return 'กำลังดำเนินการ';
    case 'WAITING':     return 'รอข้อมูลเพิ่มเติม';
    case 'CLOSED':      return 'ปิดเรื่องแล้ว';
    case 'REJECTED':    return 'ไม่รับดำเนินการ';
    default:            return status || 'ไม่ระบุ';
  }
}

/** ดึงไฟล์แนบจาก e.files ตามชื่อ field "attachment" */
function getAttachmentBlobs_(e) {
  const blobs = [];
  if (!e.files) return blobs;

  const field = e.files['attachment'];
  if (!field) return blobs;

  if (Array.isArray(field)) {
    field.forEach((b) => blobs.push(b));
  } else {
    blobs.push(field);
  }
  return blobs;
}

// ========== HANDLERS: ประชาชน ==========

function handleCreateCitizenCase(e) {
  try {
    const ss    = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_CITIZEN);

    const params = e.parameter;
    const now    = new Date();

    // เช็คช่องจำเป็น
    if (!params.full_name || !params.phone || !params.email ||
        !params.case_type || !params.case_detail || !params.line_id) {
      return jsonResponse({ error: 'missing_required_fields' });
    }
    if (!params.pdpa) {
      return jsonResponse({ error: 'pdpa_not_accepted' });
    }

    const newId = Utilities.getUuid();

    const header   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    CITIZEN_HEADER.forEach((name) => {
      colIndex[name] = header.indexOf(name) + 1; // 1-based
    });

    const lastRow = sheet.getLastRow();
    const seq     = lastRow; // ใช้จำนวนแถวเป็น running number

    const referenceNo =
      'C-' +
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd') +
      '-' +
      Utilities.formatString('%04d', seq);

    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_CITIZEN);
    const caseFolder = rootFolder.createFolder(referenceNo);

    const blobs = getAttachmentBlobs_(e);
    blobs.forEach((blob) => {
      caseFolder.createFile(blob);
    });

    const pdpaConsent = params.pdpa ? 'AGREE' : '';

    const rowIndex = lastRow + 1;
    sheet.getRange(rowIndex, colIndex['id']).setValue(newId);
    sheet.getRange(rowIndex, colIndex['reference_no']).setValue(referenceNo);
    sheet.getRange(rowIndex, colIndex['created_at']).setValue(now);
    sheet.getRange(rowIndex, colIndex['citizen_name']).setValue(params.full_name || '');
    sheet.getRange(rowIndex, colIndex['phone']).setValue(params.phone || '');
    sheet.getRange(rowIndex, colIndex['line_id']).setValue(params.line_id || '');
    sheet.getRange(rowIndex, colIndex['email']).setValue(params.email || '');
    sheet.getRange(rowIndex, colIndex['case_type']).setValue(params.case_type || '');
    sheet.getRange(rowIndex, colIndex['case_detail']).setValue(params.case_detail || '');
    sheet.getRange(rowIndex, colIndex['status']).setValue('RECEIVED');
    sheet.getRange(rowIndex, colIndex['drive_folder_id']).setValue(caseFolder.getId());
    sheet.getRange(rowIndex, colIndex['last_update']).setValue(now);
    sheet.getRange(rowIndex, colIndex['pdpa_consent']).setValue(pdpaConsent);
    sheet.getRange(rowIndex, colIndex['admin_note']).setValue('');

    appendLog_('PUBLIC', 'CREATE_CITIZEN_CASE', 'CITIZEN', referenceNo, 'สร้างคำร้องใหม่จากประชาชน');

    // ส่งอีเมลแจ้งเตือน (ถ้าตั้งค่า MailApp ได้)
    if (params.email) {
      try {
        MailApp.sendEmail({
          to: params.email,
          subject: 'สภาทนายความภาค 7 - ระบบได้รับคำร้องของท่านแล้ว',
          htmlBody:
            'เรียนคุณ ' + (params.full_name || '') +
            '<br><br>ระบบได้บันทึกคำร้องของท่านแล้ว เลขอ้างอิงคือ <b>' + referenceNo +
            '</b><br>กรุณาเก็บเลขนี้ไว้เพื่อติดตามสถานะในภายหลัง'
        });
      } catch (mailErr) {
        appendLog_('SYSTEM', 'MAIL_FAIL', 'CITIZEN', referenceNo, mailErr.toString());
      }
    }

    return jsonResponse({
      reference_no: referenceNo,
      status: 'ok'
    });

  } catch (err) {
    Logger.log('handleCreateCitizenCase error: ' + err);
    appendLog_('SYSTEM', 'ERROR_CREATE_CITIZEN', 'CITIZEN', '', err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

function handleGetCitizenStatus(e) {
  try {
    const ref   = (e.parameter.reference_no || '').trim();
    const email = (e.parameter.email || '').trim().toLowerCase();

    if (!ref || !email) {
      return jsonResponse({ error: 'missing_parameters' });
    }

    const ss    = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_CITIZEN);
    const data  = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return jsonResponse({ error: 'not_found' });
    }

    const header   = data[0];
    const idxRef   = header.indexOf('reference_no');
    const idxEmail = header.indexOf('email');
    const idxStatus= header.indexOf('status');
    const idxLast  = header.indexOf('last_update');

    for (let i = 1; i < data.length; i++) {
      const row      = data[i];
      const rowRef   = String(row[idxRef]   || '').trim();
      const rowEmail = String(row[idxEmail] || '').trim().toLowerCase();

      if (rowRef === ref && rowEmail === email) {
        const status     = row[idxStatus];
        const lastUpdate = row[idxLast];
        const statusTh   = mapStatusToThai_(status);

        return jsonResponse({
          reference_no: ref,
          status: status,
          status_th: statusTh,
          last_update: lastUpdate
            ? Utilities.formatDate(lastUpdate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
            : ''
        });
      }
    }

    return jsonResponse({ error: 'not_found' });

  } catch (err) {
    Logger.log('handleGetCitizenStatus error: ' + err);
    appendLog_('SYSTEM', 'ERROR_GET_STATUS_CITIZEN', 'CITIZEN', '', err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

// ========== HANDLERS: ทนายความ ==========

function handleCreateLawyerHelpCase(e) {
  try {
    const ss    = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_LAWYER);

    const params = e.parameter;
    const now    = new Date();

    if (!params.full_name ||
        !params.lawyer_license ||
        !params.region ||
        !params.phone ||
        !params.line_id ||
        !params.email ||
        !params.issue_type ||
        !params.issue_detail) {
      return jsonResponse({ error: 'missing_required_fields' });
    }
    if (!params.pdpa) {
      return jsonResponse({ error: 'pdpa_not_accepted' });
    }

    const newId = Utilities.getUuid();

    const header   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = {};
    LAWYER_HEADER.forEach((name) => {
      colIndex[name] = header.indexOf(name) + 1;
    });

    const lastRow    = sheet.getLastRow();
    const seq        = lastRow;
    const referenceNo =
      'L-' +
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd') +
      '-' +
      Utilities.formatString('%04d', seq);

    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_LAWYER);
    const caseFolder = rootFolder.createFolder(referenceNo);

    const blobs = getAttachmentBlobs_(e);
    blobs.forEach((blob) => {
      caseFolder.createFile(blob);
    });

    const pdpaConsent = params.pdpa ? 'AGREE' : '';

    const rowIndex = lastRow + 1;
    sheet.getRange(rowIndex, colIndex['id']).setValue(newId);
    sheet.getRange(rowIndex, colIndex['reference_no']).setValue(referenceNo);
    sheet.getRange(rowIndex, colIndex['created_at']).setValue(now);
    sheet.getRange(rowIndex, colIndex['lawyer_name']).setValue(params.full_name || '');
    sheet.getRange(rowIndex, colIndex['lawyer_license']).setValue(params.lawyer_license || '');
    sheet.getRange(rowIndex, colIndex['region']).setValue(params.region || '');
    sheet.getRange(rowIndex, colIndex['phone']).setValue(params.phone || '');
    sheet.getRange(rowIndex, colIndex['line_id']).setValue(params.line_id || '');
    sheet.getRange(rowIndex, colIndex['email']).setValue(params.email || '');
    sheet.getRange(rowIndex, colIndex['issue_type']).setValue(params.issue_type || '');
    sheet.getRange(rowIndex, colIndex['issue_detail']).setValue(params.issue_detail || '');
    sheet.getRange(rowIndex, colIndex['status']).setValue('RECEIVED');
    sheet.getRange(rowIndex, colIndex['drive_folder_id']).setValue(caseFolder.getId());
    sheet.getRange(rowIndex, colIndex['last_update']).setValue(now);
    sheet.getRange(rowIndex, colIndex['pdpa_consent']).setValue(pdpaConsent);
    sheet.getRange(rowIndex, colIndex['admin_note']).setValue('');

    appendLog_('LAWYER', 'CREATE_LAWYER_CASE', 'LAWYER', referenceNo, 'สร้างคำขอช่วยเหลือจากทนายความ');

    if (params.email) {
      try {
        MailApp.sendEmail({
          to: params.email,
          subject: 'สภาทนายความภาค 7 - ระบบได้รับคำขอช่วยเหลือของท่านแล้ว',
          htmlBody:
            'เรียนคุณ ' + (params.full_name || '') +
            '<br><br>ระบบได้บันทึกคำขอช่วยเหลือของท่านแล้ว เลขอ้างอิงคือ <b>' + referenceNo +
            '</b><br>กรุณาเก็บเลขนี้ไว้เพื่อติดตามสถานะในภายหลัง'
        });
      } catch (mailErr) {
        appendLog_('SYSTEM', 'MAIL_FAIL', 'LAWYER', referenceNo, mailErr.toString());
      }
    }

    return jsonResponse({
      reference_no: referenceNo,
      status: 'ok'
    });

  } catch (err) {
    Logger.log('handleCreateLawyerHelpCase error: ' + err);
    appendLog_('SYSTEM', 'ERROR_CREATE_LAWYER', 'LAWYER', '', err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

function handleGetLawyerHelpStatus(e) {
  try {
    const ref   = (e.parameter.reference_no || '').trim();
    const email = (e.parameter.email || '').trim().toLowerCase();

    if (!ref || !email) {
      return jsonResponse({ error: 'missing_parameters' });
    }

    const ss    = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_LAWYER);
    const data  = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return jsonResponse({ error: 'not_found' });
    }

    const header   = data[0];
    const idxRef   = header.indexOf('reference_no');
    const idxEmail = header.indexOf('email');
    const idxStatus= header.indexOf('status');
    const idxLast  = header.indexOf('last_update');

    for (let i = 1; i < data.length; i++) {
      const row      = data[i];
      const rowRef   = String(row[idxRef]   || '').trim();
      const rowEmail = String(row[idxEmail] || '').trim().toLowerCase();

      if (rowRef === ref && rowEmail === email) {
        const status     = row[idxStatus];
        const lastUpdate = row[idxLast];
        const statusTh   = mapStatusToThai_(status);

        return jsonResponse({
          reference_no: ref,
          status: status,
          status_th: statusTh,
          last_update: lastUpdate
            ? Utilities.formatDate(lastUpdate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
            : ''
        });
      }
    }

    return jsonResponse({ error: 'not_found' });

  } catch (err) {
    Logger.log('handleGetLawyerHelpStatus error: ' + err);
    appendLog_('SYSTEM', 'ERROR_GET_STATUS_LAWYER', 'LAWYER', '', err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

// ========== HANDLERS: ADMIN DASHBOARD ==========

/**
 * adminSummary
 *  - สรุปจำนวนเรื่องประชาชน / ทนาย และจำนวนสถานะรวม
 */
function handleAdminSummary(e) {
  const ss = getSpreadsheet_();

  const summary = {
    citizen_total: 0,
    lawyer_total: 0,
    total: 0,
    status_all: {
      RECEIVED: 0,
      IN_PROGRESS: 0,
      WAITING: 0,
      CLOSED: 0,
      REJECTED: 0
    }
  };

  // CITIZEN
  const citizenSheet = ss.getSheetByName(SHEET_CITIZEN);
  if (citizenSheet) {
    const data = citizenSheet.getDataRange().getValues();
    if (data.length > 1) {
      const header = data[0];
      const idxStatus = header.indexOf('status');
      for (let i = 1; i < data.length; i++) {
        const status = String(data[i][idxStatus] || '').trim();
        summary.citizen_total++;
        summary.total++;
        if (summary.status_all[status] !== undefined) {
          summary.status_all[status]++;
        }
      }
    }
  }

  // LAWYER
  const lawyerSheet = ss.getSheetByName(SHEET_LAWYER);
  if (lawyerSheet) {
    const data = lawyerSheet.getDataRange().getValues();
    if (data.length > 1) {
      const header = data[0];
      const idxStatus = header.indexOf('status');
      for (let i = 1; i < data.length; i++) {
        const status = String(data[i][idxStatus] || '').trim();
        summary.lawyer_total++;
        summary.total++;
        if (summary.status_all[status] !== undefined) {
          summary.status_all[status]++;
        }
      }
    }
  }

  return jsonResponse(summary);
}

/**
 * adminListCases
 *  module=citizen|lawyer
 *  status=(optional) ALL/RECEIVED/IN_PROGRESS/...
 */
function handleAdminListCases(e) {
  const module = (e.parameter.module || 'citizen').toLowerCase();
  const statusFilter = (e.parameter.status || 'ALL').toUpperCase();

  const ss = getSpreadsheet_();
  const sheetName = module === 'lawyer' ? SHEET_LAWYER : SHEET_CITIZEN;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ error: 'sheet_not_found' });

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse({ module, items: [] });

  const header = data[0];
  const idxRef    = header.indexOf('reference_no');
  const idxName   = header.indexOf(module === 'lawyer' ? 'lawyer_name' : 'citizen_name');
  const idxType   = header.indexOf(module === 'lawyer' ? 'issue_type'   : 'case_type');
  const idxStatus = header.indexOf('status');
  const idxCreated= header.indexOf('created_at');
  const idxLast   = header.indexOf('last_update');

  const items = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[idxStatus] || '').toUpperCase();
    if (statusFilter !== 'ALL' && status !== statusFilter) continue;

    items.push({
      reference_no: row[idxRef],
      name:        row[idxName],
      type:        row[idxType],
      status:      status,
      status_th:   mapStatusToThai_(status),
      created_at:  row[idxCreated]
        ? Utilities.formatDate(row[idxCreated], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        : '',
      last_update: row[idxLast]
        ? Utilities.formatDate(row[idxLast], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        : ''
    });
  }

  return jsonResponse({ module, items });
}

/**
 * adminUpdateStatus
 *  module=citizen|lawyer
 *  reference_no=...
 *  status=RECEIVED/IN_PROGRESS/WAITING/CLOSED/REJECTED
 *  admin_note=(optional)
 */
function handleAdminUpdateStatus(e) {
  const module = (e.parameter.module || 'citizen').toLowerCase();
  const ref    = (e.parameter.reference_no || '').trim();
  const status = (e.parameter.status || '').toUpperCase();
  const note   = (e.parameter.admin_note || '').trim();

  if (!ref || !status) {
    return jsonResponse({ error: 'missing_parameters' });
  }

  const allowed = ['RECEIVED', 'IN_PROGRESS', 'WAITING', 'CLOSED', 'REJECTED'];
  if (allowed.indexOf(status) === -1) {
    return jsonResponse({ error: 'invalid_status' });
  }

  const ss = getSpreadsheet_();
  const sheetName = module === 'lawyer' ? SHEET_LAWYER : SHEET_CITIZEN;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ error: 'sheet_not_found' });

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse({ error: 'not_found' });

  const header   = data[0];
  const idxRef   = header.indexOf('reference_no');
  const idxStatus= header.indexOf('status');
  const idxLast  = header.indexOf('last_update');
  const idxNote  = header.indexOf('admin_note');

  let foundRow = -1;

  for (let i = 1; i < data.length; i++) {
    const rowRef = String(data[i][idxRef] || '').trim();
    if (rowRef === ref) {
      foundRow = i + 1; // 1-based
      break;
    }
  }

  if (foundRow === -1) {
    return jsonResponse({ error: 'not_found' });
  }

  const now = new Date();
  sheet.getRange(foundRow, idxStatus + 1).setValue(status);
  sheet.getRange(foundRow, idxLast + 1).setValue(now);
  if (idxNote >= 0) {
    sheet.getRange(foundRow, idxNote + 1).setValue(note);
  }

  appendLog_('ADMIN', 'ADMIN_UPDATE_STATUS', module.toUpperCase(), ref, 'เปลี่ยนสถานะเป็น ' + status + ' : ' + note);

  return jsonResponse({
    ok: true,
    reference_no: ref,
    status,
    status_th: mapStatusToThai_(status),
    last_update: Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  });
}

// ========== ฟังก์ชัน setup เริ่มต้น ==========

/**
 * ใช้รันจาก Script Editor ครั้งแรกเพื่อเช็คว่าเชื่อม Sheet/Drive ได้
 * และให้โครงสร้างชีตถูกต้อง
 */
function initSystem() {
  ensureSheetsStructure_();

  const folderCitizen = DriveApp.getFolderById(DRIVE_ROOT_CITIZEN);
  const folderLawyer  = DriveApp.getFolderById(DRIVE_ROOT_LAWYER);

  Logger.log('Citizen root folder: ' + folderCitizen.getName());
  Logger.log('Lawyer root folder: ' + folderLawyer.getName());

  // เพิ่ม return ไว้เผื่อเรียกผ่านเว็บได้ด้วย (ไม่กระทบการเรียกจาก Editor)
  return jsonResponse({ ok: true, message: 'initSystem completed' });
}
