# ระบบบันทึกสุขภาพและกิจวัตรประจำวันนักเรียน — Vercel Edition

แพ็กเกจนี้ปรับหน้าเว็บเดิมให้ Deploy บน Vercel ได้ โดยคง Google Apps Script (GAS) เป็น Backend และ Google Sheets เป็นฐานข้อมูลเดิมไว้ ไม่ต้องย้ายข้อมูลนักเรียนไปฐานข้อมูลใหม่

## โครงสร้าง

- `index.html` — หน้าเว็บเดิม ปรับการเรียก API ให้ผ่าน `/api/rpc`
- `api/rpc.js` — Vercel Serverless Function ทำหน้าที่เป็น Proxy ไปยัง Google Apps Script
- `vercel.json` — กำหนดเวลาทำงานของ API สูงสุด 60 วินาที
- `code(9).gs` — โค้ด Backend เดิม (ต้องติดตั้ง/อัปเดตใน Apps Script ของ Google Sheet เดิม)

## ขั้นตอนที่ 1: ติดตั้ง Backend (Google Apps Script)

1. เปิด Google Sheet ที่ใช้เก็บข้อมูลเดิม
2. ไปที่ **Extensions > Apps Script**
3. นำโค้ดจากไฟล์ `code(9).gs` ไปแทน/รวมในโปรเจกต์ Apps Script เดิม โดยต้องรักษา Script Properties และข้อมูลชีตเดิมไว้
4. ตรวจสอบว่ามีฟังก์ชัน `doPost(e)` อยู่ท้ายไฟล์ (ไฟล์ที่แนบมามีอยู่แล้ว)
5. กด Save
6. ไปที่ **Deploy > New deployment**
7. เลือกประเภท **Web app**
8. ตั้งค่า:
   - Execute as: **Me**
   - Who has access: **Anyone**
9. กด Deploy และอนุญาตสิทธิ์ตามขั้นตอน
10. คัดลอก URL ที่ลงท้ายด้วย `/exec`

> สำคัญ: การตั้ง Anyone จำเป็นเพื่อให้ Vercel ส่งคำขอเข้ามาได้ ระบบสิทธิ์ใช้งานจริงยังตรวจ Username/PIN และ Session ที่ Backend ทุกครั้ง ห้ามนำ URL นี้ไปใส่ใน HTML หรือเปิดเผยใน Git repository ให้เก็บไว้ใน Environment Variable ของ Vercel เท่านั้น

### การอัปเดตโค้ด GAS ในอนาคต

เมื่อแก้โค้ด Apps Script ให้ไปที่ **Deploy > Manage deployments** แล้วแก้ไขแต่ละ Deployment ที่ต้องการอัปเดต โดยเลือก New version จากนั้น Deploy อีกครั้ง อย่าสลับการตั้งค่า Execute as ของ Deployment Admin กับ Deployment สำหรับ Vercel

## ขั้นตอนที่ 2: Deploy หน้าเว็บบน Vercel

1. แตก ZIP นี้
2. สร้าง Git repository ใหม่ แล้วใส่ไฟล์/โฟลเดอร์ทั้งหมดใน ZIP ลงไปที่ Root ของ repository
3. Push ขึ้น GitHub
4. เข้า https://vercel.com แล้วเลือก **Add New > Project**
5. Import repository ที่สร้างไว้
6. Framework Preset เลือก **Other**
7. Build Command และ Output Directory ปล่อยว่างได้ (เป็น Static HTML)
8. ก่อน Deploy หรือหลัง Import ให้เข้า **Project Settings > Environment Variables**
9. เพิ่มตัวแปร:

   - Name: `GAS_WEB_APP_URL`
   - Value: URL `/exec` ที่คัดลอกจาก Apps Script

10. เลือก Environment ให้ครบตามที่ใช้งาน เช่น Production, Preview, Development
11. กด Deploy

หลัง Deploy แล้วเปิด URL ของ Vercel เพื่อใช้งาน

## ข้อจำกัดที่ต้องทราบ

- ข้อมูลยังเก็บใน Google Sheets เดิม ไม่ได้ย้ายไป Vercel Database
- ครูและผู้อำนวยการเข้าใช้งานผ่าน Vercel ด้วย Username/PIN ได้
- Admin เข้าใช้งานจาก Vercel ได้ด้วย Username + Password แยกจากบัญชีครู/ผู้อำนวยการ (ครูยังใช้ PIN 6 หลักเหมือนเดิม)
- ก่อนใช้ Admin ครั้งแรก ให้เปิด Google Sheet ต้นฉบับ แล้วเลือกเมนู **ระบบบันทึกสุขภาพ > ตั้ง Username / Password สำหรับ Admin (ใช้บน Vercel)** เมนูนี้อนุญาตเฉพาะ Google Admin ที่ตั้งไว้เดิมเท่านั้น
- ถ้ายังไม่เคยตั้งค่าความปลอดภัย ให้เลือก **1) ตั้งค่าความปลอดภัย / ผู้ดูแลระบบ (ทำครั้งแรก)** ก่อน แล้วจึงตั้ง Username / Password ของ Admin
- Password ของ Admin จะถูกเก็บเป็นค่าแฮชใน Script Properties ไม่ได้เก็บเป็นข้อความปกติ การเปลี่ยน Password จะทำให้ Session Admin เดิมหมดอายุ
- ยังควรเก็บ Deployment เดิมของ Google Apps Script ไว้สำหรับจัดการโครงการ/ชีตและเมนูผู้ดูแลระบบ ทั้ง Deployment เดิมและ Deployment สำหรับ Vercel ใช้ Apps Script project และ Google Sheet ชุดเดียวกัน
- หากต้องการย้าย Backend ออกจาก Apps Script ทั้งหมดในอนาคต ต้องพัฒนาระบบใหม่ เช่น Vercel Functions + PostgreSQL/Supabase และย้ายข้อมูล/ระบบสิทธิ์อย่างเป็นขั้นตอน

## ตั้งค่าบัญชี Admin ก่อนเข้า Vercel ครั้งแรก

1. เปิด Google Sheet ที่ผูกกับ Apps Script และเปิดเมนู **ระบบบันทึกสุขภาพ**
2. หากยังไม่เคยตั้งระบบ ให้กด **1) ตั้งค่าความปลอดภัย / ผู้ดูแลระบบ (ทำครั้งแรก)** และทำตามหน้าต่างที่แสดง
3. กด **ตั้ง Username / Password สำหรับ Admin (ใช้บน Vercel)**
4. ตั้ง Username 3-30 ตัวอักษร และ Password อย่างน้อย 10 ตัวอักษร จากนั้นยืนยัน Password
5. กลับไปหน้า Vercel แล้วเข้าสู่ระบบในส่วน **ผู้ดูแลระบบ (Admin)** ด้วยข้อมูลที่ตั้งไว้

อย่าใช้ Password ซ้ำกับอีเมลหรือบัญชีสำคัญอื่น และอย่าส่ง Password ผ่านกลุ่มแชต

## ตรวจสอบเมื่อ Deploy แล้วมีปัญหา

- ขึ้น `ยังไม่ได้ตั้งค่า GAS_WEB_APP_URL`: ตรวจ Environment Variable และ Redeploy
- ขึ้น `Google Apps Script ตอบกลับไม่ใช่ JSON`: ตรวจ URL `/exec`, Deploy access เป็น Anyone และสิทธิ์การอนุญาตของ Apps Script
- ขึ้น `เชื่อมต่อ Google Apps Script ไม่สำเร็จ`: ตรวจว่า Deployment ยัง Active และ URL ถูกต้อง
- เข้าระบบได้แต่ไม่มีห้อง/ข้อมูล: ตรวจว่า Apps Script ผูกกับ Google Sheet ไฟล์เดิม ไม่ใช่ Spreadsheet เปล่า

## ความปลอดภัย

ระบบมีข้อมูลสุขภาพนักเรียนซึ่งเป็นข้อมูลส่วนบุคคล ควรจำกัดผู้ที่เข้าถึง URL, ใช้ PIN ที่ไม่ซ้ำ/เดายาก, ไม่แชร์ URL ของ Google Apps Script โดยไม่จำเป็น และเปิดใช้การป้องกันบัญชี Google ของเจ้าของข้อมูล
