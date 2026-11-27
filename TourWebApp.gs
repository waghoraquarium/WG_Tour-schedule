/**
 * ไฟล์เสริมสำหรับเว็บแอป ตารางนำชม_หว้ากอ (วางเพิ่มเป็นไฟล์ใหม่ชื่อ TourWebApp.gs)
 * ใช้งานร่วมกับสคริปต์ LINE Bot เดิมได้โดยไม่ทับฟังก์ชันเดิม
 *
 * ขั้นตอน:
 * 1) ใน Apps Script กดปุ่ม + เพิ่มไฟล์ Script ใหม่ชื่อ TourWebApp แล้ววางโค้ดนี้ลงไป
 *    (หากใช้ "ดึงข้อมูลอย่างเดียว" ไม่ต้องแก้โค้ดเดิมเลย เพราะของเดิมยังไม่มี doGet)
 * 2) หากต้องการให้เว็บแอป "เพิ่ม/แก้ไข/ลบ" ลงชีตได้ด้วย ให้แทรกบรรทัดนี้ไว้บนสุดใน doPost(e) เดิม:
 *    if (json && json.app === 'waghor_tour_web') return handleTourWebPost(json);
 * 3) (แนะนำเพื่อความปลอดภัย) ตั้งค่า Script Properties ชื่อ TOUR_API_TOKEN แล้วกรอก Token เดียวกันในหน้าเว็บแอป
 * 4) กด Deploy > Manage deployments > แก้ไขเป็น New version แล้วนำ URL มาวางที่นี่
 */
const TOUR_SPREADSHEET_ID = '14ADP9vp_OcTGnNszw8bNUptOQ6w1_EwaZMprUguqLYA';

function findTourSheet_(ss, sheetName) {
  const raw = String(sheetName || '').trim();
  if (!raw) return null;
  let sheet = ss.getSheetByName(raw);
  if (!sheet) {
    const targetNorm = raw.replace(/\s+/g, '');
    sheet = ss.getSheets().find(s => s.getName().replace(/\s+/g, '') === targetNorm) || null;
  }
  return sheet;
}

function getScheduleSheetNames_(ss) {
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => !/^(เจ้าหน้าที่|template|config|สรุป)/i.test(String(name || '').trim()));
}

function extractDayNumber_(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    return val.getDate();
  }
  const n = Number(val);
  return Number.isFinite(n) ? n : 0;
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(TOUR_SPREADSHEET_ID);
    const sheetName = (e && e.parameter && e.parameter.sheet) ? e.parameter.sheet : 'พ.ย.69';
    const sheet = findTourSheet_(ss, sheetName);
    if (!sheet) return tourJsonResponse({ ok: false, error: 'ไม่พบแท็บชีต: ' + sheetName });

    const lastRow = sheet.getLastRow();
    const rows = [];
    const sheets = getScheduleSheetNames_(ss);

    if (sheet.getName().trim() === '2/69') {
      if (lastRow < 11) {
        return tourJsonResponse({ ok: true, sheet: sheetName, rows: [], staff: readStaffRoster_(ss), sheets: sheets });
      }
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
      if (lastRow < 4) {
        return tourJsonResponse({ ok: true, sheet: sheetName, rows: [], staff: readStaffRoster_(ss), sheets: sheets });
      }
      const values = sheet.getRange(4, 1, lastRow - 3, 11).getValues();
      let lastDay = 1;
      values.forEach((r, idx) => {
        const dNum = extractDayNumber_(r[0]);
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
    return tourJsonResponse({ ok: true, sheet: sheetName, rows: rows, staff: staff, sheets: sheets });
  } catch (err) {
    return tourJsonResponse({ ok: false, error: err.message });
  }
}

function handleTourWebPost(body) {
  try {
    const props = typeof PropertiesService !== 'undefined' ? PropertiesService.getScriptProperties() : null;
    const expectedToken = props ? String(props.getProperty('TOUR_API_TOKEN') || '').trim() : '';
    if (expectedToken && String((body && body.token) || '').trim() !== expectedToken) {
      return tourJsonResponse({ ok: false, error: 'ไม่ได้รับอนุญาต (API Token ไม่ถูกต้อง)' });
    }
    if (!body || !body.record || typeof body.record !== 'object') {
      return tourJsonResponse({ ok: false, error: 'ข้อมูลที่ส่งมาไม่ถูกต้อง' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(TOUR_SPREADSHEET_ID);
    const sheet = findTourSheet_(ss, body.sheet);
    if (!sheet) return tourJsonResponse({ ok: false, error: 'ไม่พบแท็บชีต: ' + body.sheet });

    const r = body.record;
    const isLegacy269 = sheet.getName().trim() === '2/69';
    const startRow = isLegacy269 ? 11 : 4;
    const lastRow = sheet.getLastRow();

    const legacyRowValues = [
      Number(r.day) || '',
      tourSanitizeCell(r.org),
      tourSanitizeCell(r.province),
      Number(r.ft) || '',
      Number(r.fs) || '',
      Number(r.it) || '',
      Number(r.is) || '',
      Number(r.gen) || '',
      tourSanitizeCell(r.assignee)
    ];

    const standardRowValues = [
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

    if (body.action === 'delete') {
      const rowNum = Number(r.rowNumber);
      if (!Number.isInteger(rowNum) || rowNum < startRow || rowNum > lastRow) {
        return tourJsonResponse({ ok: false, error: 'เลขแถวไม่ถูกต้องสำหรับการลบ (กรุณาดึงข้อมูลล่าสุดจากชีตก่อน)' });
      }
      if (isLegacy269) {
        sheet.getRange(rowNum, 6, 1, 9).deleteCells(SpreadsheetApp.Dimension.ROWS);
      } else {
        if (rowNum < lastRow) {
          const deletedDayRaw = sheet.getRange(rowNum, 1).getValue();
          const nextDayRaw = sheet.getRange(rowNum + 1, 1).getValue();
          const nextOrg = String(sheet.getRange(rowNum + 1, 3).getValue() || '').trim();
          const nextDayNum = extractDayNumber_(nextDayRaw);
          if (nextOrg && !(nextDayNum >= 1 && nextDayNum <= 31)) {
            const carryDay = extractDayNumber_(deletedDayRaw) || Number(r.day) || '';
            if (carryDay >= 1 && carryDay <= 31) {
              sheet.getRange(rowNum + 1, 1).setValue(carryDay);
            }
          }
        }
        sheet.deleteRow(rowNum);
      }
      return tourJsonResponse({ ok: true, action: 'deleted', rowNumber: rowNum });
    }

    if (body.action === 'update') {
      const rowNum = Number(r.rowNumber);
      if (!Number.isInteger(rowNum) || rowNum < startRow || rowNum > lastRow) {
        return tourJsonResponse({ ok: false, error: 'เลขแถวไม่ถูกต้องสำหรับการแก้ไข (กรุณาดึงข้อมูลล่าสุดจากชีตก่อน)' });
      }
      if (isLegacy269) {
        sheet.getRange(rowNum, 6, 1, 9).setValues([legacyRowValues]);
      } else {
        sheet.getRange(rowNum, 1, 1, 11).setValues([standardRowValues]);
      }
      return tourJsonResponse({ ok: true, action: 'updated', rowNumber: rowNum });
    }

    if (body.action === 'create') {
      if (isLegacy269) {
        const nextRow = Math.max(11, lastRow + 1);
        sheet.getRange(nextRow, 6, 1, 9).setValues([legacyRowValues]);
        return tourJsonResponse({ ok: true, action: 'created', rowNumber: nextRow });
      }
      sheet.appendRow(standardRowValues);
      return tourJsonResponse({ ok: true, action: 'created', rowNumber: sheet.getLastRow() });
    }

    return tourJsonResponse({ ok: false, error: 'ไม่รองรับคำสั่ง: ' + String(body.action || '') });
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