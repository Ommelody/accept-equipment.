/**
 * ระบบตรวจรับครุภัณฑ์ | SAWARiN
 * -----------------------------------------------------------
 * Backend (Google Apps Script)
 * เชื่อมฐานข้อมูลเดิม: Spreadsheet ที่ผูกกับสคริปต์ + โฟลเดอร์ Drive เดิม
 * ** ไม่มีการแก้ไข logic หรือการเชื่อมต่อฐานข้อมูล **
 * -----------------------------------------------------------
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบตรวจรับครุภัณฑ์ | SAWARiN')
    .setFaviconUrl("https://img.icons8.com/fluency/48/add-camera.png")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ดึงข้อมูลเริ่มต้น
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  const masterSheet = ss.getSheetByName('Masterครุภัณฑ์');
  
  if (!settingsSheet || !masterSheet) throw new Error('ไม่พบชีท Settings หรือ Masterครุภัณฑ์');

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

// อัปเดต/นำเข้าข้อมูล PO
function updatePOData(dataArray) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let poSheet = ss.getSheetByName('Data_PO');
    if (!poSheet) poSheet = ss.insertSheet('Data_PO');
    
    poSheet.clearContents(); 
    if(dataArray.length > 0) {
      poSheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
    }
    return { success: true, message: 'อัปเดตข้อมูล PO เรียบร้อยแล้ว' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ค้นหาข้อมูล PO
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
      qty: r[17], // ดึงจำนวนจากคอลัมน์ R (Index 17)
      unit: r[19], 
      price: String(r[22]).replace(/,/g, '').trim()      
    }));

    return { success: true, poDate: results[0][6], vendor: results[0][10], items: items };
  } else {
    return { success: false, message: 'ไม่พบเลขที่ PO นี้ในระบบ' };
  }
}

// บันทึกข้อมูล
function saveFormData(formObject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Data_Log');
    if(!sheet) sheet = ss.insertSheet('Data_Log');

    const folderId = "1F--VGstbziNKIwrWA45VUylSprd0Wv3a"; 
    const folder = DriveApp.getFolderById(folderId);
    const items = JSON.parse(formObject.items_json);
    
    // สำคัญ: สร้าง Timestamp ครั้งเดียว เพื่อให้ทุกไอเทมในล็อตนี้มีเวลาตรงกันเป๊ะ
    const batchTimestamp = new Date().toLocaleString('th-TH'); 
    
    items.forEach(item => {
      let imageUrl = "ไม่มีรูปภาพ";
      
      if (item.image) {
        const content = Utilities.base64Decode(item.image.split(",")[1]);
        const blob = Utilities.newBlob(content, "image/jpeg", "ASSET_" + new Date().getTime() + "_" + Math.floor(Math.random()*1000) + ".jpg");
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        imageUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
      }

      sheet.appendRow([
        batchTimestamp, // ใช้เวลาเดียวกันในการจัดกลุ่ม
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

// โหลดประวัติแบบมี Limit และจัดกลุ่ม (Group by Timestamp)
function getHistory(limit) {
  limit = parseInt(limit) || 50;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Log');
  if (!sheet) return { success: false, message: "ไม่พบชีท Data_Log" };
  
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return { success: true, groups: [] };

  const headers = data.shift();
  
  // จัดกลุ่มข้อมูลโดยใช้ Timestamp (คอลัมน์ A / index 0) เป็น Key
  let groupedData = {};
  
  data.forEach(row => {
    let timestamp = row[0];
    if (!groupedData[timestamp]) {
      groupedData[timestamp] = {
        timestamp: row[0],
        po_no: row[1],
        po_date: row[2],
        receive_no: row[3],
        receive_date: row[4],
        vendor: row[5],
        items: [] // เก็บ array ของครุภัณฑ์ในบิลนี้
      };
    }
    groupedData[timestamp].items.push({
      assetId: row[6], itemName: row[7], unit: row[8], price: row[9],
      type: row[10], brand: row[11], model: row[12], sn: row[13],
      dept: row[14], floor: row[15], room: row[16], image: row[17]
    });
  });

  // แปลง Object เป็น Array แล้ว Reverse เอาใบใหม่สุดขึ้นก่อน
  let groupArray = Object.values(groupedData).reverse().slice(0, limit);

  return { success: true, groups: groupArray }; 
}

// แก้ไขรายการครุภัณฑ์ 1 ชิ้นในใบตรวจรับ (ระบุด้วย timestamp ของบิล + ลำดับภายในบิล)
function updateHistoryItem(timestamp, itemIndex, itemData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Log');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Data_Log' };

    // ใช้ getDisplayValues ให้ค่า timestamp ตรงกับที่ getHistory ส่งไปหน้าเว็บเป๊ะ ๆ
    const data = sheet.getDataRange().getDisplayValues();
    let matchRows = [];
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() === String(timestamp).trim()) matchRows.push(r + 1); // เลขแถวจริงในชีท (1-indexed)
    }
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

    // คอลัมน์ G..Q (assetId..room) = 7 ถึง 17 รวม 11 คอลัมน์, คอลัมน์ R (image) = 18
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

// ลบรายการครุภัณฑ์ 1 ชิ้นออกจากใบตรวจรับ (ระบุด้วย timestamp ของบิล + ลำดับภายในบิล)
function deleteHistoryItem(timestamp, itemIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data_Log');
    if (!sheet) return { success: false, message: 'ไม่พบชีท Data_Log' };

    // ใช้ getDisplayValues ให้ค่า timestamp ตรงกับที่ getHistory ส่งไปหน้าเว็บเป๊ะ ๆ
    const data = sheet.getDataRange().getDisplayValues();
    let matchRows = [];
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() === String(timestamp).trim()) matchRows.push(r + 1);
    }
    if (itemIndex < 0 || itemIndex >= matchRows.length) return { success: false, message: 'ไม่พบรายการนี้ในระบบ' };

    sheet.deleteRow(matchRows[itemIndex]);
    return { success: true, message: 'ลบรายการเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
