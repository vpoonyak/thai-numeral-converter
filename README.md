# IT๙ Converter (Word Add-in)

เครื่องมือแปลงเลขไทยสำหรับ Microsoft Word ที่มาพร้อมกับระบบ "Smart Ignore" เพื่อความปลอดภัยในการใช้งานร่วมกับเอกสารที่มีภาษาอังกฤษและ URL

## ฟีเจอร์หลัก (Features)
- **Smart Ignore (ข้ามคำอังกฤษอัตโนมัติ)**: ระบบจะข้ามตัวเลขที่อยู่ในคำภาษาอังกฤษ (เช่น spin9, iPhone 15) และ URL/Email โดยอัตโนมัติ
- **Deep Search & Flatten**: ปุ่มพิเศษสำหรับแปลงและ "แช่แข็ง" องค์ประกอบที่เข้าถึงยาก:
  - ส่วนหัวและท้ายกระดาษ (Headers & Footers)
  - เลขหน้าและคำบรรยายภาพ (Fields & Page Numbers)
  - ลำดับรายการอัตโนมัติ (Auto-lists 1.1 -> ๑.๑)
- **Formatting Preservation**: รักษาฟอนต์ สี และขนาดดั้งเดิมไว้ครบถ้วน

## 📥 ดาวน์โหลด (Download manifest.xml)
*เพื่อให้ไฟล์โหลดลงเครื่องได้โดยตรง กรุณา **คลิกขวา** ที่ลิงก์ด้านล่างแล้วเลือก **"Save Link As..." (บันทึกลิงก์เป็น...)***
👉 [**ดาวน์โหลดไฟล์ manifest.xml**](https://vitchakorn.com/thai-numeral-converter/manifest.xml)

## เอกสารทางกฎหมาย (Legal Documents)
สำหรับการตรวจสอบโดย Microsoft Partner Center:
- **Privacy Policy**: [Link](https://vitchakorn.com/thai-numeral-converter/privacy.html)
- **End User License Agreement (EULA)**: [Link](https://vitchakorn.com/thai-numeral-converter/eula.html)

## 💻 วิธีการติดตั้ง (How to Install)

### สำหรับ MacOS
เพื่อให้ปุ่ม **IT๙** ปรากฏในโปรแกรม Word บน Mac ให้ทำตามขั้นตอนดังนี้:

1. **Download & Copy**: [ดาวน์โหลดไฟล์ `manifest.xml`](https://vitchakorn.com/thai-numeral-converter/manifest.xml) (*คลิกขวา -> Save As*) แล้วคัดลอกไปไว้ที่:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml`
   *(หากไม่มีโฟลเดอร์ `wef` ให้สร้างขึ้นมาใหม่)*
2. **Restart Word**: ปิดโปรแกรม Word (Cmd + Q) แล้วเปิดใหม่
3. **เปิดใช้งาน Add-in**:
   - ไปที่แถบ **หน้าแรก (Home)**
   - มองหาปุ่ม **Add-ins** ทางด้านขวา (หรือไปที่ **Insert** > **Add-ins** > **My Add-ins**)
   - เลือกหัวข้อ **DEVELOPER ADD-INS**
   - จะพบ **IT๙ Converter** ปรากฏอยู่ ให้กดเพิ่ม (Add)

### สำหรับ Windows
การติดตั้งบน Windows จะใช้วิธีการทำ **Shared Folder (Sideloading)**:

1. **เตรียมโฟลเดอร์**: สร้างโฟลเดอร์ใหม่ (เช่น `C:\Addins`) และนำไฟล์ `manifest.xml` ไปวางไว้ในนั้น
2. **แชร์โฟลเดอร์**:
   - คลิกขวาที่โฟลเดอร์ -> **Properties** -> แถบ **Sharing** -> กดปุ่ม **Share**
   - เลือกชื่อผู้ใช้ของคุณเองแล้วกด **Share**
   - คัดลอก **Network Path** ที่ได้ (เช่น `\\ชื่อคอมพิวเตอร์\Addins`)
3. **ตั้งค่าใน Word**:
   - เปิด Word -> ไปที่ **File (ไฟล์)** -> **Options (ตัวเลือก)** -> **Trust Center (ศูนย์ความเชื่อถือ)**
   - กดปุ่ม **Trust Center Settings (การตั้งค่าศูนย์ความเชื่อถือ)** -> **Trusted Add-in Catalogs (แค็ตตาล็อก Add-in ที่เชื่อถือได้)**
   - ในช่อง **Catalog URL** ให้วาง Network Path ที่คัดลอกมา แล้วกด **Add Catalog (เพิ่มแค็ตตาล็อก)**
   - ติ๊กถูกที่ช่อง **Show in Menu (แสดงในเมนู)** ของรายการนั้น แล้วกด OK และ Restart Word
4. **เปิดใช้งาน**:
   - ไปที่ **Insert (แทรก)** -> **My Add-ins (Add-in ของฉัน)**
   - เลือกแถบ **SHARED FOLDER (โฟลเดอร์ที่แชร์)** จะพบ **IT๙ Converter** ให้กด Add

## การตั้งค่า
- **แปลงเนื้อหาหลัก**: แปลงเฉพาะตัวเลขในส่วนเนื้อหาหลักของเอกสาร
- **แปลงรวมเลขหน้าและเลขหัวข้อแบบ Flatten**: แปลงและ flatten เลขในหัวกระดาษ, เลขหน้า และลำดับรายการอัตโนมัติ (จะกลายเป็นข้อความถาวร)

---
**พัฒนาโดย:** Vitchakorn Poonyakanok (IT๙)
**เวอร์ชัน:** 1.23.0
