/**
 * ระบบตรวจรับครุภัณฑ์ | SAWARiN — Backend (Google Apps Script Web App)
 * -----------------------------------------------------------
 * Frontend อยู่บน GitHub Pages เรียกผ่าน HTTP (fetch)
 * Backend นี้ทำหน้าที่เป็น API เชื่อม Google Sheets + Drive (ฐานข้อมูลเดิม)
 *
 * วิธี Deploy:
 *   Deploy > New deployment > Web app
 *   Execute as: Me   |   Who has access: Anyone
 *   คัดลอก Web app URL ไปใส่ในหน้าตั้งค่าของแอป (ไอคอนเฟือง)
 * -----------------------------------------------------------
 */

// ===== Router =====
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const fn = API[req.action];
    if (!fn) return json_({ success: false, message: 'ไม่รู้จักคำสั่ง: ' + req.action });
    const result = fn.apply(null, req.args || []);
    return json_(result);
  } catch (err) {
    return json_({ success: false, message: 'Server error: ' + err });
  }
}

// รองรับ health-check / JSONP (?action=...&args=...&callback=...)
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) {
    return json_({ ok: true, service: 'ASSET System API', time: new Date().toISOString() });
  }
  try {
    const fn = API[p.action];
    if (!fn) return jsonp_(p.callback, { success: false, message: 'ไม่รู้จักคำสั่ง' });
    const args = p.args ? JSON.parse(p.args) : [];
    const result = fn.apply(null, args);
    return jsonp_(p.callback, result);
  } catch (err) {
    return jsonp_(p.callback, { success: false, message: 'Server error: ' + err });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function jsonp_(cb, obj) {
  if (!cb) return json_(obj);
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(obj) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ===== รายการ API ที่เปิดให้ frontend เรียก =====
const API = {
  getInitialData: getInitialData,
  updatePOData: updatePOData,
  searchPO: searchPO,
  saveFormData: saveFormData,
  getHistory: getHistory,
  updateHistoryItem: updateHistoryItem,
  deleteHistoryItem: deleteHistoryItem,
  importPORows: importPORows,
  logPOImport: logPOImport,
  getImportStatus: getImportStatus,
  getReportData: getReportData,
  listPOsByDate: listPOsByDate,
  generateFixedAssetNumbers: generateFixedAssetNumbers,
  importRunningMax: importRunningMax
};

// ===== Logic (เชื่อมฐานข้อมูลเดิม) =====

function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  const masterSheet = ss.getSheetByName('Masterครุภัณฑ์');
  if (!settingsSheet || !masterSheet) return { success: false, message: 'ไม่พบชีท Settings หรือ Masterครุภัณฑ์' };

  const settings = settingsSheet.getDataRange().getValues();
  const master = masterSheet.getDataRange().getValues();
  settings.shift();
  master.shift();

  return {
    vendors: [...new Set(settings.map(r => r[0]).filter(String))],
    depts: [...new Set(settings.map(r => r[1]).filter(String))],
    floors: [...new Set(settings.map(r => r[2]).filter(String))],
    rooms: [...new Set(settings.map(r => r[3]).filter(String))],
    types: [...new Set(settings.map(r => r[4]).filter(String))],
    masterData: master.map(r => ({
      type: String(r[0]),
      id: String(r[1]),
      name: String(r[2]),
      unit: String(r[18])
    })).filter(m => m.id)
  };
}

function updatePOData(dataArray) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let poSheet = ss.getSheetByName('Data_PO');
    if (!poSheet) poSheet = ss.insertSheet('Data_PO');
    poSheet.clearContents();
    if (dataArray.length > 0) {
      poSheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
    }
    return { success: true, message: 'อัปเดตข้อมูล PO เรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// นำเข้า PO แบบต่อท้าย (ข้ามแถวซ้ำ: Document Number + No.) — ข้อมูลเก็บเป็นข้อความเพื่อคงรูปแบบวันที่/รหัส
function importPORows(rows, header) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('Data_PO');
    if (!sh) sh = ss.insertSheet('Data_PO');
    let last = sh.getLastRow();
    const W = 36;
    if (last === 0 && header && header.length) {
      sh.getRange(1, 1, 1, W).setNumberFormat('@').setValues([header.concat(new Array(W).fill('')).slice(0, W)]);
      last = 1;
    }
    const seen = {};
    if (last >= 1) {
      const d = sh.getRange(1, 4, last, 1).getDisplayValues();
      const n = sh.getRange(1, 15, last, 1).getDisplayValues();
      for (let i = 0; i < last; i++) seen[d[i][0] + '|' + n[i][0]] = 1;
    }
    const add = []; let skipped = 0;
    rows.forEach(function (r) {
      const k = r[3] + '|' + r[14];
      if (seen[k]) { skipped++; return; }
      seen[k] = 1;
      add.push(r.concat(new Array(W).fill('')).slice(0, W).map(String));
    });
    if (add.length) {
      sh.getRange(last + 1, 1, add.length, W).setNumberFormat('@').setValues(add);
    }
    return { success: true, added: add.length, skipped: skipped };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// บันทึกประวัติการอัปโหลด (วัน/เวลา, ไฟล์, ช่วง Posting Date)
function logPOImport(info) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('Import_Log');
    if (!sh) {
      sh = ss.insertSheet('Import_Log');
      sh.getRange(1, 1, 1, 8).setValues([['เวลาอัปโหลด', 'ไฟล์', 'อ่านทั้งหมด', 'เข้าเงื่อนไข', 'เพิ่มใหม่', 'ข้ามซ้ำ', 'Posting เริ่ม', 'Posting ล่าสุด']]);
    }
    const at = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 8).setNumberFormat('@').setValues([[at, info.file, info.read, info.kept, info.added, info.skipped, info.min, info.max]]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getImportStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Import_Log');
    const po = ss.getSheetByName('Data_PO');
    const totalRows = po ? Math.max(0, po.getLastRow() - 1) : 0;
    if (!sh || sh.getLastRow() < 2) return { success: true, last: null, totalRows: totalRows };
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getDisplayValues();
    let maxPosting = '';
    data.forEach(function (r) { if (r[7] && r[7] > maxPosting) maxPosting = r[7]; });
    const hist = data.slice(-5).map(function (r) { return { at: r[0], file: r[1], added: r[4], min: r[6], max: r[7] }; });
    return { success: true, last: { at: data[data.length - 1][0], file: data[data.length - 1][1] }, maxPosting: maxPosting, totalRows: totalRows, history: hist };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getReportData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Log');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Data_Log' };
    const data = sheet.getDataRange().getDisplayValues();
    data.shift();
    return { success: true, rows: data.filter(function (r) { return r[0] !== '' || r[1] !== ''; }).map(function (r) { return r.slice(0, 19); }) };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function searchPO(poNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poSheet = ss.getSheetByName('Data_PO');
  if (!poSheet) return { success: false, message: 'ไม่พบชีท Data_PO กรุณานำเข้าข้อมูลก่อน' };

  const data = poSheet.getDataRange().getDisplayValues();
  const results = data.filter(r => r[3] === poNumber);

  if (results.length > 0) {
    const items = results.map(r => ({
      assetId: r[15],
      itemName: r[16],
      qty: r[17],
      unit: r[19],
      price: String(r[22]).replace(/,/g, '').trim(),
      dept: String(r[30] || '').trim()
    }));
    return { success: true, poDate: results[0][6], vendor: results[0][10], items: items };
  }
  return { success: false, message: 'ไม่พบเลขที่ PO นี้ในระบบ' };
}

// รายการ PO ตามวันที่ (คอลัมน์ G = วันที่ในเอกสาร) — dateStr รูปแบบ yyyy-mm-dd
function parseDateLoose_(s) {
  s = String(s || '').trim(); if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/), d, mo, y;
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else { m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/); if (!m) return null; d = +m[1]; mo = +m[2]; y = +m[3]; }
  if (y < 100) y += 2000; if (y > 2400) y -= 543;
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}

function listPOsByDate(dateStr) {
  try {
    const poSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_PO');
    if (!poSheet) return { success: false, message: 'ไม่พบชีท Data_PO กรุณานำเข้าข้อมูลก่อน' };
    const data = poSheet.getDataRange().getDisplayValues();
    const map = {}, order = [];
    for (var i = 1; i < data.length; i++) {
      if (parseDateLoose_(data[i][6]) !== dateStr) continue;
      var po = data[i][3]; if (!po) continue;
      if (!map[po]) { map[po] = { po: po, vendor: data[i][10], count: 0 }; order.push(po); }
      map[po].count++;
    }
    return { success: true, pos: order.map(function (k) { return map[k]; }) };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function saveFormData(formObject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Data_Log');
    if (!sheet) sheet = ss.insertSheet('Data_Log');

    const folderId = "1F--VGstbziNKIwrWA45VUylSprd0Wv3a";
    const folder = DriveApp.getFolderById(folderId);
    const items = JSON.parse(formObject.items_json);

    const batchTimestamp = new Date().toLocaleString('th-TH');

    items.forEach(item => {
      let imageUrl = "ไม่มีรูปภาพ";
      if (item.image) {
        const content = Utilities.base64Decode(item.image.split(",")[1]);
        const blob = Utilities.newBlob(content, "image/jpeg", "ASSET_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000) + ".jpg");
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        imageUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
      }
      sheet.appendRow([
        batchTimestamp,
        formObject.po_no, formObject.po_date, formObject.receive_no, formObject.receive_date, formObject.vendor,
        item.assetId, item.itemName, item.unit, item.price, item.type, item.brand, item.model, item.sn,
        item.dept, item.floor, item.room, imageUrl, item.faNo || ''
      ]);
    });
    return { success: true, message: "บันทึกข้อมูลเรียบร้อยแล้ว" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getHistory(limit) {
  limit = parseInt(limit) || 50;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Log');
  if (!sheet) return { success: false, message: "ไม่พบชีท Data_Log" };

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return { success: true, groups: [] };
  data.shift();

  let groupedData = {};
  data.forEach(row => {
    let timestamp = row[0];
    if (!groupedData[timestamp]) {
      groupedData[timestamp] = {
        timestamp: row[0], po_no: row[1], po_date: row[2],
        receive_no: row[3], receive_date: row[4], vendor: row[5], items: []
      };
    }
    groupedData[timestamp].items.push({
      assetId: row[6], itemName: row[7], unit: row[8], price: row[9],
      type: row[10], brand: row[11], model: row[12], sn: row[13],
      dept: row[14], floor: row[15], room: row[16], image: row[17], faNo: row[18] || ''
    });
  });

  let groupArray = Object.values(groupedData).reverse().slice(0, limit);
  return { success: true, groups: groupArray };
}

// หาแถวทั้งหมดของ "บิล" เดียวกัน โดยระบุด้วยชุดคีย์ (timestamp / po_no / receive_no)
function findGroupRows_(data, groupKey) {
  var ts  = String((groupKey && groupKey.timestamp)  || '').trim();
  var po  = String((groupKey && groupKey.po_no)      || '').trim();
  var rec = String((groupKey && groupKey.receive_no) || '').trim();

  var byAll = [], byTs = [], byPoRec = [];
  for (var r = 1; r < data.length; r++) {
    var rowTs  = String(data[r][0]).trim();
    var rowPo  = String(data[r][1]).trim();
    var rowRec = String(data[r][3]).trim();
    var tsMatch  = ts !== '' && rowTs === ts;
    var poMatch  = po !== '' && rowPo === po;
    var recMatch = rowRec === rec;
    if (tsMatch && poMatch && recMatch) byAll.push(r + 1);
    if (tsMatch) byTs.push(r + 1);
    if (poMatch && recMatch) byPoRec.push(r + 1);
  }
  if (byAll.length) return byAll;
  if (byTs.length) return byTs;
  return byPoRec;
}

function updateHistoryItem(groupKey, itemIndex, itemData) {
  try {
    if (typeof groupKey === 'string') groupKey = { timestamp: groupKey };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Log');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Data_Log' };

    const data = sheet.getDataRange().getDisplayValues();
    const matchRows = findGroupRows_(data, groupKey);
    if (itemIndex < 0 || itemIndex >= matchRows.length) return { success: false, message: 'ไม่พบรายการนี้ในระบบ' };
    const rowNum = matchRows[itemIndex];

    let imageUrl = itemData.existingImage || '';
    if (itemData.removeImage) {
      imageUrl = 'ไม่มีรูปภาพ';
    } else if (itemData.image) {
      const folderId = "1F--VGstbziNKIwrWA45VUylSprd0Wv3a";
      const folder = DriveApp.getFolderById(folderId);
      const content = Utilities.base64Decode(itemData.image.split(",")[1]);
      const blob = Utilities.newBlob(content, "image/jpeg", "ASSET_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000) + ".jpg");
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    }

    sheet.getRange(rowNum, 7, 1, 11).setValues([[
      itemData.assetId, itemData.itemName, itemData.unit, itemData.price,
      itemData.type, itemData.brand, itemData.model, itemData.sn,
      itemData.dept, itemData.floor, itemData.room
    ]]);
    sheet.getRange(rowNum, 18).setValue(imageUrl);
    sheet.getRange(rowNum, 19).setNumberFormat('@').setValue(itemData.faNo || '');

    return { success: true, message: 'แก้ไขข้อมูลเรียบร้อยแล้ว', imageUrl: imageUrl };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteHistoryItem(groupKey, itemIndex) {
  try {
    if (typeof groupKey === 'string') groupKey = { timestamp: groupKey };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Log');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Data_Log' };

    const data = sheet.getDataRange().getDisplayValues();
    const matchRows = findGroupRows_(data, groupKey);
    if (itemIndex < 0 || itemIndex >= matchRows.length) return { success: false, message: 'ไม่พบรายการนี้ในระบบ' };

    sheet.deleteRow(matchRows[itemIndex]);
    return { success: true, message: 'ลบรายการเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}


// ===== Generate เลขครุภัณฑ์ (Fix Asset) : YY-<รหัสพัสดุ 10 หลัก>-<รัน 4 หลัก> =====
// ค่าสูงสุดของแต่ละกลุ่ม (ปี + รหัสพัสดุ) หาจากชีท SAP_Running (นำเข้าจากไฟล์ SAP) และ Data_Log คอลัมน์ S
function getRunningMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = {};
  const put = function (prefix, n) { if (!(map[prefix] >= n)) map[prefix] = n; };
  const sr = ss.getSheetByName('SAP_Running');
  if (sr && sr.getLastRow() >= 2) {
    sr.getRange(2, 1, sr.getLastRow() - 1, 2).getDisplayValues().forEach(function (r) { put(String(r[0]).trim(), parseInt(r[1], 10) || 0); });
  }
  const dl = ss.getSheetByName('Data_Log');
  if (dl && dl.getLastRow() >= 2) {
    dl.getRange(2, 19, dl.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) {
      const m = /^(\d{2}-\d{10})-(\d{4})$/.exec(String(r[0]).trim());
      if (m) put(m[1], parseInt(m[2], 10));
    });
  }
  return map;
}

function generateFixedAssetNumbers(req) {
  try {
    const yy = String(req.year || '').replace(/\D/g, '').slice(-2);
    if (yy.length !== 2) return { success: false, message: 'ปีไม่ถูกต้อง' };
    const map = getRunningMap_();
    const numbers = [], warnings = [];
    (req.ids || []).forEach(function (id, i) {
      const mat = String(id || '').replace(/\D/g, '');
      if (mat.length !== 10) { numbers.push(''); warnings.push('บรรทัดที่ ' + (i + 1) + ': รหัสพัสดุ "' + (id || '') + '" ไม่ครบ 10 หลัก'); return; }
      const prefix = yy + '-' + mat;
      const next = (map[prefix] || 0) + 1;
      if (next > 9999) { numbers.push(''); warnings.push('บรรทัดที่ ' + (i + 1) + ': รันเกิน 9999'); return; }
      map[prefix] = next;
      numbers.push(prefix + '-' + ('0000' + next).slice(-4));
    });
    return { success: true, numbers: numbers, warnings: warnings };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// นำเข้าเลขรันนิ่งสูงสุดต่อกลุ่มจากไฟล์ SAP: list = [{p:'24-8010030002', max:12}]
function importRunningMax(list) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('SAP_Running');
    if (!sh) { sh = ss.insertSheet('SAP_Running'); sh.getRange(1, 1, 1, 3).setValues([['กลุ่ม (ปี-รหัสพัสดุ)', 'เลขรันสูงสุด', 'อัปเดตล่าสุด']]); }
    const idx = {};
    if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(function (r, i) { idx[String(r[0]).trim()] = { row: i + 2, max: parseInt(r[1], 10) || 0 }; });
    const at = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    const add = []; let updated = 0;
    list.forEach(function (x) {
      const cur = idx[x.p];
      if (cur) { if (x.max > cur.max) { sh.getRange(cur.row, 2, 1, 2).setNumberFormat('@').setValues([[String(x.max), at]]); updated++; } }
      else add.push([x.p, String(x.max), at]);
    });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setNumberFormat('@').setValues(add);
    return { success: true, added: add.length, updated: updated };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
