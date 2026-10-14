# คู่มือใช้งาน HRMI AUTO

## ภาพรวม

ระบบแบ่งเป็น 3 flow แยกจากกัน:

1. ประมวลผลเวลาการทำงานใน Prosoft HRMI
2. Export รายงาน `ขาด` และ `สาย` เป็น Excel
3. Import ไฟล์ Excel เข้า Rocket75 ให้ตรงกับข้อมูลขาดสะสม ข้อมูลสายสะสม และข้อมูลลาสะสม

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

## ติดตั้งบนเครื่องใหม่

หลังติดตั้ง Prosoft HRMI, Python 3.10 ขึ้นไป และ Node.js LTS ให้เปิด PowerShell ในโฟลเดอร์โปรเจกต์แล้วรัน:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install-new-machine.ps1
```

สามารถดับเบิลคลิก `install-new-machine.cmd` แทนคำสั่งข้างต้นได้

สคริปต์จะสร้าง `.venv`, ติดตั้ง dependencies, สร้าง `.env` จากตัวอย่างเมื่อยังไม่มี และตรวจโปรแกรมที่จำเป็น จากนั้นกรอกบัญชีใน `.env` และปรับเทียบตำแหน่งเมนูของเครื่องใหม่:

```powershell
.\.venv\Scripts\python.exe .\hrmi.py --auto-login --stop-before-next --calibrate-time-attendance
.\install-new-machine.ps1 -InstallDailyTask
```

คำสั่งสุดท้ายจะสร้าง Scheduled Task เวลา 10:00 ใช้ `-DailyAt "07:30"` เมื่อต้องการเปลี่ยนเวลา งานใช้สิทธิ์ผู้ใช้ปกติเมื่อเปิด PowerShell ปกติ และใช้สิทธิ์สูงสุดเมื่อเปิดแบบ **Run as administrator** ให้สิทธิ์ตรงกับ HRMI ที่ใช้งาน

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
| `ลา` | นำเข้าข้อมูลลาสะสม |

สคริปต์ Login ด้วย `ROCKET_USER` และ `ROCKET_PASSWORD` และสร้าง key/token ใหม่ในหน่วยความจำทุกครั้ง token จะไม่ถูกบันทึกลงไฟล์

ทุกประเภท Import ตรวจเฉพาะวันที่ Created ของโฟลเดอร์ที่เก็บไฟล์ ซึ่งต้องเป็นวันปัจจุบันตามเวลาไทย ไม่ตรวจ Created หรือ Modified ของแต่ละไฟล์ และไม่ใช้วันที่ในชื่อโฟลเดอร์แทน Created โฟลเดอร์ที่สร้างวันอื่นจะถูกข้าม แม้ระบุช่วงวันที่เอง

ก่อน Import ระบบจะอ่านทุก worksheet หากไฟล์มีเฉพาะหัวตารางและไม่มีข้อมูลตั้งแต่แถวที่ 2 ระบบจะแสดง `ข้ามไฟล์เปล่า` และไม่ส่งไฟล์นั้นไป Rocket75 หากเว็บแจ้งข้อมูลซ้ำหรือข้อผิดพลาด ระบบจะหยุดเพื่อให้ตรวจสอบก่อนรันต่อ

Import รายงานอนุมัติลาจาก Share Center:

```powershell
.\run-rocket75-import.ps1 -Types "leave" -LeaveRoot "T:\HR_Leave_Approved"
```

คำสั่งนี้เลือกเฉพาะโฟลเดอร์ย่อยที่มี `CreationTime` เป็นวันที่ปัจจุบัน หากไม่พบหรือพบมากกว่า 1 โฟลเดอร์ ระบบจะหยุดเพื่อป้องกันการ Import ผิดรอบ หากต้องการระบุโฟลเดอร์เอง ให้ใช้ `-LeaveFolder "T:\HR_Leave_Approved\101069"`

## ติดตั้ง Scheduled Task

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\setup-daily-task.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\setup-export-import-tasks.ps1
```

| เวลาไทยทุกวัน | Scheduled Task | งาน |
|---|---|---|
| 10:00 | `HRMI AUTO Daily` | ประมวลผลเวลา HRMI |
| 14:00 | `HRMI AUTO Export` | Export ขาดและสาย |
| 14:30 | `HRMI AUTO Rocket75 Import` | Import ขาดและสายเข้า Rocket75 |
| 15:00 | `HRMI AUTO Leave Import` | Import ลาจาก `T:\HR_Leave_Approved` |

คำสั่งติดตั้งจะสร้างหรืออัปเดตงานชื่อเดิม ไม่เริ่ม Import ทันที Windows ต้องใช้เขตเวลา Bangkok (`SE Asia Standard Time`) ต้องเปิดเครื่อง Login Windows ค้างไว้ ไม่ Lock หน้าจอ และไม่ให้เครื่อง Sleep ระหว่างงาน HRMI รวมถึงต้องเข้าถึงไดรฟ์ `T:` ได้ตอน Import ลา งาน Export และ Import ไม่ลองใหม่อัตโนมัติ งานประมวลผลเวลาลองใหม่ทุก 30 นาที สูงสุด 3 ครั้งเมื่อผิดพลาด

เวลาของแต่ละงานเป็นตารางแยกกัน ไม่ได้รอให้งานก่อนหน้าจบ ควรตรวจให้ Export เสร็จก่อนรอบ Import

ตรวจสถานะและรอบถัดไป:

```powershell
Get-ScheduledTask -TaskName 'HRMI AUTO*' | Get-ScheduledTaskInfo
```

Log ของงานอัตโนมัติ:

- ประมวลผลเวลา: `logs\scheduler.log`
- Export: `logs\export-scheduler.log`
- Import ขาดและสาย: `logs\import-scheduler.log`
- Import ลา: `logs\leave-import-scheduler.log`

เรียกสคริปต์ Import ลาพร้อมบันทึก log ด้วยตนเอง:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\run-rocket75-leave-scheduled.ps1
```

เปลี่ยน Share Center root ได้ด้วย `-LeaveRoot "เส้นทางโฟลเดอร์"` สคริปต์หยุดเมื่อเข้าถึงโฟลเดอร์ไม่ได้ หรือไม่พบโฟลเดอร์ย่อย Created วันนี้เพียงหนึ่งโฟลเดอร์

ตัวสร้าง Excel บันทึกเวลาใช้ Python `openpyxl` จาก `requirements.txt` ไม่ต้องติดตั้ง `@oai/artifact-tool`

## ตรวจสอบก่อนใช้งาน

```powershell
.\.venv\Scripts\python.exe -m unittest -v test_hrmi_schedule.py
```

ควรตรวจไฟล์ Excel และ Log หลังจบรอบ โดยเฉพาะเมื่อ HRMI มีการเปลี่ยนหน้าตาหรืออัปเดตเวอร์ชัน
