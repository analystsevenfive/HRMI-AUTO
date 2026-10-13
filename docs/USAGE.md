# คู่มือใช้งาน HRMI AUTO

## ภาพรวม

ระบบแบ่งเป็น 3 flow แยกจากกัน:

1. ประมวลผลเวลาการทำงานใน Prosoft HRMI
2. Export รายงาน `ขาด` และ `สาย` เป็น Excel
3. Import ไฟล์ Excel เข้า Rocket75 ให้ตรงกับข้อมูลขาดสะสมและข้อมูลสายสะสม

ลำดับสาขาคือ `[BB]`, `[BK]`, `[NTS]`, `[PK]`, `[SM]`

## การตั้งค่า `.env`

สร้าง `.env` ที่ root ของโปรเจกต์:

```env
HRMI_USER=
HRMI_PASSWORD=
ROCKET_USER=
ROCKET_PASSWORD=
```

ไฟล์นี้ถูกยกเว้นจาก Git ห้าม commit ชื่อผู้ใช้ รหัสผ่าน key หรือ token

## กฎช่วงวันที่อัตโนมัติ

- วันอังคาร: ใช้วันเสาร์ถึงวันอาทิตย์ล่าสุด
- วันอื่น: ใช้วันเดียวย้อนหลัง 2 วัน
- การระบุ `-StartDate` และ `-EndDate` จะใช้ช่วงที่กำหนดแทน

## 1. ประมวลผลเวลา

รันอัตโนมัติตามกฎวันที่:

```powershell
.\run-hrmi-scheduled.ps1
```

กำหนดช่วงวันที่เอง:

```powershell
.\run-hrmi-manual.ps1 -StartDate 10/10/2569 -EndDate 11/10/2569
```

ผลการทำงานและระยะเวลาอยู่ใน `logs/` และ `HRMI-processing-times.xlsx`

## 2. Export รายงานขาดและสาย

```powershell
.\run-hrmi-reports.ps1
```

กำหนดช่วงและสั่ง Export ทับ:

```powershell
.\run-hrmi-reports.ps1 -StartDate 10/10/2569 -EndDate 11/10/2569 -Force
```

ระบบตั้งค่ารายงานดังนี้:

- ประเภทรายงาน: `แบบแจกแจง(นับสายเช้า-บ่าย)`
- ประเภทข้อมูล: `ขาด` หรือ `สาย` ตามรอบ
- ช่วงวันที่: ช่วงเดียวกับการประมวลผล
- รายงานข้อมูลการอนุมัติลา: ปิดชั่วคราว

ไฟล์อยู่ใน `exports\<ช่วงวันที่>\` และเวลาการ Export อยู่ใน `HRMI-report-export-times.xlsx`

## 3. Import เข้า Rocket75

Import ทั้งขาดและสายครบทุกสาขา:

```powershell
.\run-rocket75-import.ps1
```

กำหนดช่วงวันที่:

```powershell
.\run-rocket75-import.ps1 -StartDate 10/10/2569 -EndDate 11/10/2569
```

เลือกเฉพาะประเภทหรือสาขา:

```powershell
.\run-rocket75-import.ps1 -Types "absence" -Organizations "BB","SM"
.\run-rocket75-import.ps1 -Types "late" -Organizations "BB","BK","NTS","PK","SM"
```

การจับคู่ endpoint:

| ประเภทไฟล์ | ปลายทาง Rocket75 |
|---|---|
| `ขาด` | นำเข้าข้อมูลขาดสะสม |
| `สาย` | นำเข้าข้อมูลสายสะสม |

สคริปต์ Login ด้วย `ROCKET_USER` และ `ROCKET_PASSWORD` และสร้าง key/token ใหม่ในหน่วยความจำทุกครั้ง token จะไม่ถูกบันทึกลงไฟล์

ก่อน Import ระบบจะอ่านทุก worksheet หากไฟล์มีเฉพาะหัวตารางและไม่มีข้อมูลตั้งแต่แถวที่ 2 ระบบจะแสดง `ข้ามไฟล์เปล่า` และไม่ส่งไฟล์นั้นไป Rocket75 หากเว็บแจ้งข้อมูลซ้ำหรือข้อผิดพลาด ระบบจะหยุดเพื่อให้ตรวจสอบก่อนรันต่อ

## ติดตั้ง Scheduled Task

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\setup-daily-task.ps1
```

Task `HRMI AUTO Daily` รันเวลา 06:00 และทำเฉพาะ flow ประมวลผลเวลา ต้องเปิดเครื่อง Login Windows ค้างไว้ ไม่ Lock หน้าจอ และไม่ให้เครื่อง Sleep

## ตรวจสอบก่อนใช้งาน

```powershell
.\.venv\Scripts\python.exe -m unittest -v test_hrmi_schedule.py
```

ควรตรวจไฟล์ Excel และ Log หลังจบรอบ โดยเฉพาะเมื่อ HRMI มีการเปลี่ยนหน้าตาหรืออัปเดตเวอร์ชัน
