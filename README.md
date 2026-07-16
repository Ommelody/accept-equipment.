# ASSET System | SAWARiN — ระบบตรวจรับครุภัณฑ์

ระบบตรวจรับครุภัณฑ์บน Google Apps Script + Google Sheets/Drive
UI ธีมมืด/สว่าง รองรับมือถือ (mobile-first) พร้อมระบบแก้ไข/ลบรายการ, แบ่งหน้าประวัติ และพิมพ์เอกสาร

## ไฟล์
- `Code.gs` — โค้ดฝั่งเซิร์ฟเวอร์ (Apps Script) เชื่อม Spreadsheet + Drive
- `Index.html` — ส่วนติดต่อผู้ใช้ (โหลดผ่าน `HtmlService`)

## ติดตั้ง
1. เปิด Google Apps Script project ที่ผูกกับ Spreadsheet ฐานข้อมูล
2. วางเนื้อหา `Code.gs` และไฟล์ HTML ชื่อ **Index**
3. Deploy เป็น Web App

## ชีทที่ใช้
- `Settings` — ตัวเลือก (บริษัท/แผนก/ชั้น/ห้อง/ประเภท)
- `Masterครุภัณฑ์` — ทะเบียนครุภัณฑ์หลัก
- `Data_PO` — ข้อมูล PO ที่นำเข้า
- `Data_Log` — บันทึกการตรวจรับ (จัดกลุ่มด้วย timestamp)

## หมายเหตุ PWA
เพิ่มลงหน้าจอโฮม (Add to Home Screen) ได้ แต่ Service Worker / โหมดออฟไลน์เต็มรูปแบบทำไม่ได้เมื่อรันผ่าน Apps Script (ข้อจำกัดของ sandbox iframe)
