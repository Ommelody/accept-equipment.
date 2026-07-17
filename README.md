# ASSET System | SAWARiN — ระบบตรวจรับครุภัณฑ์ (PWA)

Progressive Web App (ติดตั้งลงมือถือได้ · ใช้งานออฟไลน์ได้บางส่วน)
- **Frontend** — โฮสต์บน GitHub Pages (โฟลเดอร์นี้)
- **Backend** — Google Apps Script (Web App) เชื่อม Google Sheets + Drive (ฐานข้อมูลเดิม)

## โครงสร้างไฟล์
```
index.html               หน้าแอป (PWA)
manifest.webmanifest     ข้อมูลแอปสำหรับติดตั้ง
sw.js                    Service Worker (แคช/ออฟไลน์)
icons/                   ไอคอนแอป (192/512/maskable/apple/favicon)
Code.gs                  โค้ด backend — นำไปวางใน Apps Script
```

## ขั้นตอนติดตั้ง

### 1) Backend (Google Apps Script)
1. เปิด Apps Script project ที่ผูกกับ Spreadsheet ฐานข้อมูล
2. วางเนื้อหา `Code.gs` (แทนของเดิม)
3. **Deploy › New deployment › Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
4. คัดลอก **Web app URL** (ลงท้าย `/exec`)

### 2) Frontend (GitHub Pages)
1. อัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้ขึ้น repo (เช่น `Ommelody/asset-check`)
2. Settings › Pages › Source: `main` / root (หรือ `/docs`)
3. เปิดลิงก์ที่ได้ (เช่น `https://ommelody.github.io/asset-check/`)

### 3) เชื่อมต่อ
- เปิดแอป › แตะ **ไอคอนเฟือง** มุมขวาบน › วาง Web app URL › บันทึก
- (ตั้งครั้งเดียว จำค่าไว้ในเครื่อง) เว้นว่าง = โหมดตัวอย่าง

### 4) ติดตั้งลงมือถือ
- Android/Chrome: เมนู › **Add to Home screen**
- iOS/Safari: แชร์ › **เพิ่มไปยังหน้าจอโฮม**

## ฟังก์ชัน
บันทึกการตรวจรับ · ค้นหา/นำเข้า PO · ประวัติ (แบ่งหน้า) · แก้ไข/ลบรายการ · เพิ่ม/ลบรูปภาพ · พิมพ์ใบสรุปรวม & ใบเดี่ยว · โหมดสว่าง/มืด

## ชีทที่ใช้
`Settings` · `Masterครุภัณฑ์` · `Data_PO` · `Data_Log`

## หมายเหตุ
- Service Worker แคชเฉพาะหน้าแอป (shell) — ข้อมูลจาก backend ยังต้องออนไลน์
- ต้องใช้ HTTPS (GitHub Pages เป็น HTTPS อยู่แล้ว) PWA จึงทำงานเต็มรูปแบบ
