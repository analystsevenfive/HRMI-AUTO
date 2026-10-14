# HRMI Time Attendance navigation

คู่มือคำสั่งและ flow ทั้งหมดดูที่ [docs/USAGE.md](docs/USAGE.md)

สคริปต์ Python + pywinauto เปิดหน้า **Time Attendance > บันทึกข้อมูลเวลา > ประมวลผลเวลาการทำงาน** และทำงานตามลำดับ `[BB]`, `[BK]`, `[NTS]`, `[PK]`, `[SM]`

โหมดอัตโนมัติตั้งวันย้อนหลัง 2 วันตามเวลาไทย โดยปกติวันเริ่มและวันสิ้นสุดเป็นวันเดียวกัน หากรันวันอังคาร ระบบจะตั้งเป็นช่วงวันเสาร์ถึงวันอาทิตย์ล่าสุด เช่น รันวันอังคาร 13/10/2569 จะประมวลผล 10/10/2569–11/10/2569 จากนั้นกด Next ผ่านหน้าเลือกช่วงพนักงาน ตรวจวันที่บนหน้าสรุป กด Next เพื่อเริ่มประมวลผล รอจนเสร็จ เลือก No เมื่อ HRMI ถามว่าจะเปิดหน้าตรวจสอบเพิ่มเติมหรือไม่ แล้วกด Finish

หลังประมวลผล สคริปต์สามารถเปิด **Time Attendance Report > รายงานการลา ขาด สาย** เลือก `แบบแจกแจง(นับสายเช้า-บ่าย)` แล้ว Export `ขาด` และ `สาย` เป็น Excel แยกไฟล์ต่อสาขา โดยใช้ช่วงวันที่เดียวกับการประมวลผล

## ติดตั้งและใช้งาน

ใช้ Windows ที่ติดตั้ง Prosoft HRMI และมี desktop session ที่ใช้งานได้:

ดับเบิลคลิก `install-new-machine.cmd` หรือใช้คำสั่ง:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install-new-machine.ps1
```

ตัวติดตั้งจะตรวจ Python 3.10 ขึ้นไป, Node.js และ Prosoft HRMI จากนั้นสร้าง `.venv`, ติดตั้ง dependencies, สร้าง `.env` จากตัวอย่าง และตรวจว่าโมดูลพร้อมใช้งาน โดยจะไม่เขียนทับ `.env` เดิม

กำหนด `HRMI_USER` และ `HRMI_PASSWORD` ใน environment หรือไฟล์ `.env` ก่อนใช้ `--auto-login` ไฟล์ `.env` ถูกยกเว้นจาก Git อย่าเผยแพร่ไฟล์นี้

ถ้า Login เอง ต้องรันขณะ HRMI อยู่หน้าเลือกองค์กร สคริปต์จะหยุดหากเปิดค้างที่หน้าหลัก เพราะยืนยันองค์กรปัจจุบันไม่ได้

ตัวเลือกที่ใช้บ่อย:

```powershell
.\.venv\Scripts\python.exe hrmi.py --stop-before-next --organizations '[BB]' '[BK]' --calibrate-time-attendance
.\.venv\Scripts\python.exe hrmi.py --help
```

รันด้วยมือโดยกำหนดช่วงวันที่เอง ใช้ได้ทั้ง ค.ศ. และ พ.ศ. แต่วันเริ่มและวันสิ้นสุดต้องอยู่ในเดือนเดียวกัน:

```powershell
.\run-hrmi-manual.ps1
```

สคริปต์จะถามวันเริ่มและวันสิ้นสุด หรือส่งค่าเข้าไปโดยตรง:

```powershell
.\run-hrmi-manual.ps1 -StartDate 01/10/2569 -EndDate 10/10/2569
```

สามารถเรียก Python โดยตรงได้เช่นกัน:

```powershell
.\.venv\Scripts\python.exe .\hrmi.py --auto-login --start-date 01/10/2569 --end-date 10/10/2569
```

หรือใช้รูปแบบ ISO:

```powershell
.\.venv\Scripts\python.exe .\hrmi.py --auto-login --start-date 2026-10-01 --end-date 2026-10-10
```

Flow รายงานแยกจาก Flow ประมวลผลเวลา ใช้ `run-hrmi-reports.ps1` เพื่อ Export รายงาน `ขาด` และ `สาย` โดยไม่ประมวลผลเวลาใหม่ ไฟล์อยู่ที่ `exports\<ช่วงวันที่>\` ตัวอย่าง `exports\2026-10-10\HRMI_BB_ขาด_2026-10-10.xlsx`:

```powershell
.\run-hrmi-reports.ps1
```

หากไม่ระบุวันที่ ทุกประเภทของรายงานจะใช้กฎอัตโนมัติเดียวกับการประมวลผล: วันอังคารใช้ช่วงเสาร์–อาทิตย์ล่าสุด ส่วนวันอื่นใช้วันเดียวย้อนหลัง 2 วัน หากต้องการกำหนดช่วงเอง:

```powershell
.\run-hrmi-reports.ps1 -StartDate 10/10/2569 -EndDate 10/10/2569
```

Flow `รายงานข้อมูลการอนุมัติลา` ถูกปิดชั่วคราว เพิ่ม `-Force` เมื่อต้องการ Export รายงาน `ขาด` และ `สาย` ทับไฟล์เดิม ระบบบันทึกเวลาของแต่ละสาขาและแต่ละประเภทไว้ใน `HRMI-report-export-times.xlsx` พร้อมแถวสรุปเมื่อครบ 10 รายงาน

นำไฟล์ที่ Export แล้วเข้า Rocket75 โดยจับคู่ `ขาด` ไปยังข้อมูลขาดสะสม และ `สาย` ไปยังข้อมูลสายสะสม:

```powershell
.\run-rocket75-import.ps1
```

เพิ่ม `ROCKET_USER` และ `ROCKET_PASSWORD` ใน `.env` สคริปต์จะ Login และสร้าง key/token ใหม่ในหน่วยความจำทุกครั้งโดยไม่บันทึก token ลงไฟล์ หากไม่มีชื่อผู้ใช้และรหัสผ่านจึงค่อยใช้ `ROCKET75_KEY` และ `ROCKET75_TOKEN` เป็นทางเลือก สามารถกำหนดวันที่หรือเลือกประเภทได้ เช่น:

ก่อนส่ง สคริปต์จะตรวจทุก worksheet ในไฟล์ Excel หากมีเฉพาะหัวตารางและไม่มีแถวข้อมูล จะข้ามไฟล์นั้นและไม่ส่งไป Rocket75

```powershell
.\run-rocket75-import.ps1 -StartDate 10/10/2569 -EndDate 11/10/2569 -Types "absence","late"
```

นำเข้ารายงานอนุมัติลาจาก Share Center โดยให้ระบบเลือกโฟลเดอร์ย่อยที่มี `CreationTime` เป็นวันที่ปัจจุบัน:

```powershell
.\run-rocket75-import.ps1 -Types "leave" -LeaveRoot "T:\HR_Leave_Approved"
```

ระบบจะทำงานต่อเมื่อพบโฟลเดอร์ของวันนี้เพียงหนึ่งโฟลเดอร์ หากไม่พบหรือพบมากกว่าหนึ่งโฟลเดอร์จะหยุดเพื่อป้องกันการเลือกผิดรอบ หากต้องการเลือกโฟลเดอร์เอง ใช้ `-LeaveFolder "T:\HR_Leave_Approved\101069"` ข้อมูลซ้ำที่ Rocket75 ตรวจพบจะทำให้สคริปต์หยุดก่อนส่งไฟล์ถัดไป

ใช้ `--export-path` เมื่อต้องการเปลี่ยนโฟลเดอร์ปลายทาง การอัปโหลด Google Drive ยังไม่ได้เปิดใช้ใน flow นี้

หากต้องการเลือกเฉพาะบางสาขา ให้เพิ่ม `--organizations '[BB]' '[BK]'` ระบบจะตรวจการรันซ้ำแยกตามช่วงวันที่ที่กำหนด ทุก flow ที่ใช้ช่วงอัตโนมัติจะใช้กฎวันอังคารเดียวกัน

`--stop-before-next` ใช้เมื่อต้องการเปิดหน้าและตั้งวันที่เท่านั้นโดยไม่เริ่มประมวลผล

`--calibrate-time-attendance` จะใช้ตำแหน่งที่บันทึกไว้ใน `hrmi-navigation.json` ถ้ามีรูปแบบใหม่ (Version 3) หรือรอให้คลิกเมนู Time Attendance หนึ่งครั้งเพื่อบันทึกตำแหน่งสัมพัทธ์กับ Navigator ค่าเดิมรูปแบบ Version 2 จะต้องปรับเทียบใหม่

ดูผลแต่ละขั้นใน `logs/HRMI-*.csv` สคริปต์หยุดเมื่อยืนยันหน้าต่างหรือเมนูที่ต้องใช้ไม่ได้

สคริปต์จับเวลาตั้งแต่เริ่มเลือกแต่ละสาขาจนกด Finish สำเร็จ แล้วเพิ่มผลลง `HRMI-processing-times.xlsx` หลังจบแต่ละสาขา ตารางประกอบด้วยรอบการทำงาน สาขา วันที่ประมวลผล เวลาเริ่ม เวลาสิ้นสุด ระยะเวลา สถานะ และรายละเอียด เมื่อสำเร็จครบ `[BB]`, `[BK]`, `[NTS]`, `[PK]`, `[SM]` ระบบจะเพิ่มแถว `สรุปรอบ` โดยคำนวณเวลารวมตั้งแต่เริ่ม `[BB]` จน Finish `[SM]` ให้แต่ละรอบการทำงาน หากสาขาใดหยุดด้วยข้อผิดพลาด สคริปต์จะบันทึกเวลาพร้อมสถานะ `ผิดพลาด` ก่อนหยุด

## ข้อจำกัด

HRMI ใช้ custom controls บางส่วน จึงยังต้องใช้พิกัดเทียบหน้าต่างและแป้นพิมพ์บางขั้น การเลือกองค์กรยังอาศัยตำแหน่งแถวในหน้า Organization การรันซ้ำจะเริ่มประมวลผลซ้ำสำหรับวันเป้าหมาย

ไฟล์ `เปิด-HRMI-บันทึกเวลา.ps1` เป็นสคริปต์เดิม เก็บไว้สำหรับอ้างอิงระหว่างย้ายระบบ คำสั่งหลักใหม่คือ `hrmi.py`

## รันอัตโนมัติทุกวัน

รัน PowerShell ต่อไปนี้หนึ่งครั้งเพื่อสร้าง virtual environment ติดตั้ง dependencies และสร้าง Scheduled Task ชื่อ `HRMI AUTO Daily` เวลา 06:00 ทุกวัน:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\setup-daily-task.ps1
```

Task จะเรียก `run-hrmi-scheduled.ps1` ในโหมด `--scheduled` สำหรับประมวลผลเวลาเท่านั้น โดยไม่สร้างรายงาน ตรวจผลของวันที่เป้าหมายก่อนเริ่มและข้ามสาขาที่สำเร็จแล้ว เมื่อสคริปต์จบด้วยข้อผิดพลาด Task Scheduler จะลองใหม่ทุก 30 นาที สูงสุด 3 ครั้ง

ตรวจสอบหรือสั่งรันทดสอบ Task:

```powershell
Get-ScheduledTask -TaskName 'HRMI AUTO Daily' | Get-ScheduledTaskInfo
Start-ScheduledTask -TaskName 'HRMI AUTO Daily'
Get-Content .\logs\scheduler.log -Tail 100
```

ยกเลิก Task:

```powershell
Unregister-ScheduledTask -TaskName 'HRMI AUTO Daily' -Confirm:$false
```

การรันด้วยมือจะถามก่อนหากวันที่เป้าหมายประมวลผลครบแล้ว ใช้ `--force` เมื่อต้องการรันซ้ำโดยไม่ถาม:

```powershell
.\.venv\Scripts\python.exe .\hrmi.py --auto-login --force
```

เนื่องจากระบบใช้ `pywinauto` และ HRMI ต้องใช้สิทธิ์ยกระดับ Task จะตั้งเป็น `Run only when user is logged on` และ `Run with highest privileges` เครื่องต้องเปิดและ Login Windows ค้างไว้โดยไม่ Lock หน้าจอหรือเข้า Sleep ระหว่างประมวลผล

### ติดตั้งบนเครื่องใหม่

1. ติดตั้ง Prosoft HRMI, Python 3.10 ขึ้นไป และ Node.js LTS
2. Clone `https://github.com/analystsevenfive/HRMI-AUTO.git` หรือดาวน์โหลด ZIP แล้วแตกไฟล์
3. ดับเบิลคลิก `install-new-machine.cmd` หรือรัน `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install-new-machine.ps1`
4. กรอกบัญชีใน `.env`
5. รัน `.\.venv\Scripts\python.exe .\hrmi.py --auto-login --stop-before-next --calibrate-time-attendance` เพื่อสร้างตำแหน่งเมนูสำหรับเครื่องนั้น
6. รัน `.\install-new-machine.ps1 -InstallDailyTask` เพื่อสร้าง Scheduled Task เวลา 06:00

หากต้องการเวลาอื่น เช่น 07:30:

```powershell
.\install-new-machine.ps1 -InstallDailyTask -DailyAt "07:30"
```

ไฟล์ `.env` และ `hrmi-navigation.json` เป็นค่าประจำเครื่องและไม่ถูกเก็บใน Git เมื่อติดตั้งเครื่องใหม่จึงต้องสร้างและปรับเทียบอีกครั้ง
