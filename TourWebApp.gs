/**
 * ไฟล์เสริมสำหรับเว็บแอป ตารางนำชม_หว้ากอ (วางเพิ่มเป็นไฟล์ใหม่ชื่อ TourWebApp.gs)
 * ใช้งานร่วมกับสคริปต์ LINE Bot เดิมได้โดยไม่ทับฟังก์ชันเดิม
 *
 * ขั้นตอน:
 * 1) ใน Apps Script กดปุ่ม + เพิ่มไฟล์ Script ใหม่ชื่อ TourWebApp แล้ววางโค้ดนี้ลงไป
 *    (หากใช้ "ดึงข้อมูลอย่างเดียว" ไม่ต้องแก้โค้ดเดิมเลย เพราะของเดิมยังไม่มี doGet)
 * 2) หากต้องการให้เว็บแอป "เพิ่ม/แก้ไข/ลบ" ลงชีตได้ด้วย ให้แทรก 3 บรรทัดนี้ไว้บนสุดใน doPost(e) เดิม:
 *    if (json && json.app === 'waghor_tour_web') return handleTourWebPost(json);
 * 3) กด Deploy > Manage deployments > แก้ไขเป็น New version แล้วนำ URL มาวางที่นี่
 */
const TOUR_SPREADSHEET_ID = '14ADP9vp_OcTGnNszw8bNUptOQ6w1_EwaZMprUguqLYA';

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(TOUR_SPREADSHEET_ID);
    const sheetName = (e && e.parameter && e.parameter.sheet) ? e.parameter.sheet : 'พ.ย.69';
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      const targetNorm = sheetName.replace(/\s+/g, '');
      sheet = ss.getSheets().find(s => s.getName().replace(/\s+/g, '') === targetNorm) || null;
    }
    if (!sheet) return tourJsonResponse({ ok: false, error: 'ไม่พบแท็บชีต: ' + sheetName });

    const lastRow = sheet.getLastRow();
    const rows = [];

    if (sheet.getName().trim() === '2/69') {
      if (lastRow < 11) return tourJsonResponse({ ok: true, sheet: sheetName, rows: [], staff: readStaffRoster_(ss) });
      const values = sheet.getRange(11, 6, lastRow - 10, 9).getValues();
      values.forEach((r, idx) => {
        const org = String(r[1] || '').trim();
        if (!org) return;
        rows.push({
          id: sheetName + '-' + (idx + 11),
          rowNumber: idx + 11,
          day: Number(r[0]) || (idx + 1),
          time: '',
          org: org,
          province: String(r[2] || '').trim(),
          ft: Number(r[3]) || 0,
          fs: Number(r[4]) || 0,
          it: Number(r[5]) || 0,
          is: Number(r[6]) || 0,
          gen: Number(r[7]) || 0,
          speaker: 'ขอ',
          assignee: String(r[8] || '').trim()
        });
      });
    } else {
      if (lastRow < 4) return tourJsonResponse({ ok: true, sheet: sheetName, rows: [], staff: readStaffRoster_(ss) });
      const values = sheet.getRange(4, 1, lastRow - 3, 11).getValues();
      let lastDay = 1;
      values.forEach((r, idx) => {
        const dNum = Number(r[0]);
        if (dNum >= 1 && dNum <= 31) lastDay = dNum;
        const org = String(r[2] || '').trim();
        if (!org) return;
        rows.push({
          id: sheetName + '-' + (idx + 4),
          rowNumber: idx + 4,
          day: (dNum >= 1 && dNum <= 31) ? dNum : lastDay,
          time: String(r[1] || '').trim(),
          org: org,
          province: String(r[3] || '').trim(),
          ft: Number(r[4]) || 0,
          fs: Number(r[5]) || 0,
          it: Number(r[6]) || 0,
          is: Number(r[7]) || 0,
          gen: Number(r[8]) || 0,
          speaker: String(r[9] || '').indexOf('ไม่') >= 0 ? 'ไม่ขอ' : 'ขอ',
          assignee: String(r[10] || '').trim()
        });
      });
    }
    const staff = readStaffRoster_(ss);
    return tourJsonResponse({ ok: true, sheet: sheetName, rows: rows, staff: staff });
  } catch (err) {
    return tourJsonResponse({ ok: false, error: err.message });
  }
}

function handleTourWebPost(body) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(TOUR_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(body.sheet);
    if (!sheet) return tourJsonResponse({ ok: false, error: 'ไม่พบแท็บชีต: ' + body.sheet });

    const r = body.record;
    const rowValues = [
      Number(r.day) || '',
      tourSanitizeCell(r.time),
      tourSanitizeCell(r.org),
      tourSanitizeCell(r.province),
      Number(r.ft) || '',
      Number(r.fs) || '',
      Number(r.it) || '',
      Number(r.is) || '',
      Number(r.gen) || '',
      r.speaker === 'ไม่ขอ' ? 'ไม่ขอ' : 'ขอ',
      tourSanitizeCell(r.assignee)
    ];

    if (body.action === 'delete' && r.rowNumber) {
      sheet.deleteRow(Number(r.rowNumber));
      return tourJsonResponse({ ok: true, action: 'deleted' });
    }
    if (body.action === 'update' && r.rowNumber) {
      sheet.getRange(Number(r.rowNumber), 1, 1, 11).setValues([rowValues]);
      return tourJsonResponse({ ok: true, action: 'updated' });
    }
    sheet.appendRow(rowValues);
    return tourJsonResponse({ ok: true, action: 'created' });
  } catch (err) {
    return tourJsonResponse({ ok: false, error: err.message });
  }
}

function readStaffRoster_(ss) {
  const staffSheet = ss.getSheetByName('เจ้าหน้าที่');
  if (!staffSheet || staffSheet.getLastRow() < 2) return [];
  const vals = staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, 4).getValues();
  const list = [];
  vals.forEach(r => {
    const name = String(r[1] || '').trim();
    const role = String(r[2] || '').trim();
    const status = String(r[3] || '').trim();
    if (name && (!status || status.indexOf('ไม่') === -1)) {
      list.push({ name: name, role: role });
    }
  });
  return list;
}

function tourSanitizeCell(val) {
  const s = String(val || '').trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function tourJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}