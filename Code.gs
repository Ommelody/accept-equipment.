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
  deleteHistoryItem: deleteHistoryItem
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
      price: String(r[22]).replace(/,/g, '').trim()
    }));
    return { success: true, poDate: results[0][6], vendor: results[0][10], items: items };
  }
  return { success: false, message: 'ไม่พบเลขที่ PO นี้ในระบบ' };
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
        item.dept, item.floor, item.room, imageUrl
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
      dept: row[14], floor: row[15], room: row[16], image: row[17]
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
