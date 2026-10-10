"""Open HRMI time processing pages without starting calculation."""

from __future__ import annotations

import argparse
import calendar
import csv
import json
import os
import re
import subprocess
import sys
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from pywinauto import Application, Desktop, keyboard, mouse
import win32api
import win32con
import win32gui
import win32process


ROOT = Path(__file__).resolve().parent
DEFAULT_EXE = Path(r"C:\Program Files (x86)\Prosoft\Prosoft HRMI\Bin\HRMI.exe")
ORGANIZATIONS = {"[BB]": 0, "[BK]": 1, "[NTS]": 2, "[PK]": 3, "[SM]": 4}
CONFIG = ROOT / "hrmi-navigation.json"
THAI_MONTHS = ("มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม")
BANGKOK_TZ = timezone(timedelta(hours=7))


def bangkok_target_date() -> date:
    return datetime.now(BANGKOK_TZ).date() - timedelta(days=2)


def thai_date_text(value: date) -> str:
    return f"{value.day:02d} {THAI_MONTHS[value.month - 1]} {value.year + 543}"


def parse_thai_date(value: str) -> date | None:
    match = re.fullmatch(r"\s*(\d{1,2})\s+([^\s]+)\s+(\d{4})\s*", value)
    if not match or match.group(2) not in THAI_MONTHS:
        return None
    return date(int(match.group(3)) - 543, THAI_MONTHS.index(match.group(2)) + 1,
                int(match.group(1)))


def literal_keys(value: str) -> str:
    """Escape SendKeys metacharacters in credentials."""
    escapes = {"+": "{+}", "^": "{^}", "%": "{%}", "~": "{~}",
               "(": "{(}", ")": "{)}", "{": "{{}", "}": "{}}"}
    return "".join(escapes.get(char, char) for char in value)


def load_env() -> None:
    path = ROOT / ".env"
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if line.startswith("export "):
            line = line[7:].strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if not key.isidentifier():
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        os.environ[key] = value


class Hrmi:
    def __init__(self, args: argparse.Namespace):
        self.args = args
        self.pid: int | None = None
        self.org = ""
        self.step = "เริ่มต้น"
        self.target_date = bangkok_target_date()
        self.run_id = datetime.now(BANGKOK_TZ).strftime("%Y%m%d-%H%M%S")
        self.timing_json = args.log_path / "HRMI-processing-times.json"
        self.timing_xlsx = ROOT / "HRMI-processing-times.xlsx"
        args.log_path.mkdir(parents=True, exist_ok=True)
        self.log_path = args.log_path / f"HRMI-{datetime.now():%Y%m%d-%H%M%S-%f}.csv"
        self.log_file = self.log_path.open("w", encoding="utf-8-sig", newline="")
        self.writer = csv.writer(self.log_file)
        self.writer.writerow(["Timestamp", "Step", "Status", "Organization", "Message", "Details"])

    def close(self) -> None:
        self.log_file.close()

    def log(self, step: str, status: str, message: str = "", details: str = "") -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        self.writer.writerow([timestamp, step, status, self.org, message, details])
        self.log_file.flush()
        print(f"[{timestamp}] [{status}] [{step}] {self.org} {message}".strip())

    def start(self, step: str, message: str = "") -> None:
        self.step = step
        self.log(step, "START", message)

    def done(self, message: str = "") -> None:
        self.log(self.step, "OK", message)

    def record_timing(self, organization: str, started_at: datetime,
                      finished_at: datetime, status: str, message: str = "") -> None:
        rows = []
        if self.timing_json.exists():
            try:
                rows = json.loads(self.timing_json.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                rows = []
        rows.append({
            "run_id": self.run_id,
            "organization": organization,
            "target_date": self.target_date.isoformat(),
            "started_at": started_at.isoformat(),
            "finished_at": finished_at.isoformat(),
            "duration_seconds": round((finished_at - started_at).total_seconds(), 3),
            "status": status,
            "message": message,
        })
        self.timing_json.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        builder = ROOT / "build_processing_times.mjs"
        try:
            subprocess.run(
                ["node", str(builder), str(self.timing_json), str(self.timing_xlsx)],
                cwd=ROOT,
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
            )
            self.log("บันทึกเวลา Excel", "OK",
                     f"{organization}: {(finished_at - started_at).total_seconds():.1f} วินาที; {self.timing_xlsx}")
        except (OSError, subprocess.CalledProcessError) as exc:
            details = exc.stderr.strip() if isinstance(exc, subprocess.CalledProcessError) and exc.stderr else str(exc)
            self.log("บันทึกเวลา Excel", "ERROR", f"สร้าง Excel ไม่สำเร็จ: {details}")

    def windows(self) -> list:
        return [w for w in Desktop(backend="uia").windows(process=self.pid, visible_only=True)
                if w.is_visible()]

    def window(self, title: str):
        return next((w for w in self.windows() if title.casefold() in w.window_text().casefold()), None)

    def largest_window(self):
        return max(self.windows(), key=lambda w: w.rectangle().width() * w.rectangle().height(), default=None)

    def wait_window(self, title: str, timeout: float = 8):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            found = self.window(title)
            if found:
                return found
            time.sleep(.3)
        return None

    def elements(self, text: str, *, exact: bool = False, visible: bool = True) -> list:
        found = []
        wanted = "".join(text.split()).casefold()
        for window in self.windows():
            for element in [window, *window.descendants()]:
                try:
                    name = "".join(element.window_text().split()).casefold()
                    if name and ((name == wanted) if exact else (wanted in name)):
                        if not visible or element.is_visible():
                            found.append(element)
                except Exception:
                    continue
        return found

    def wait_text(self, text: str, timeout: float = 8) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if self.elements(text):
                return True
            time.sleep(.3)
        return False

    def activate(self, window) -> None:
        win32gui.ShowWindow(window.handle, win32con.SW_RESTORE)
        try:
            win32gui.SetForegroundWindow(window.handle)
        except Exception:
            try:
                window.set_focus()
            except Exception:
                pass
        time.sleep(.2)
        if win32gui.GetForegroundWindow() != window.handle:
            raise RuntimeError(f"นำหน้าต่าง {window.window_text()} มาไว้ด้านหน้าไม่ได้")

    def click_owned(self, x: int, y: int) -> None:
        handle = win32gui.WindowFromPoint((x, y))
        _, pid = win32process.GetWindowThreadProcessId(handle)
        if pid != self.pid:
            raise RuntimeError(f"ตำแหน่งคลิก ({x},{y}) ไม่ได้อยู่บนหน้าต่าง HRMI")
        mouse.click(coords=(x, y))

    def connect(self) -> None:
        if not self.args.exe.exists():
            raise RuntimeError(f"ไม่พบ HRMI.exe: {self.args.exe}")
        self.start("เปิดหรือเชื่อมต่อ HRMI")
        for window in Desktop(backend="uia").windows(visible_only=True):
            try:
                pid = window.process_id()
                import psutil
                process_path = Path(psutil.Process(pid).exe())
                if process_path.resolve().samefile(self.args.exe.resolve()):
                    self.pid = pid
                    break
            except (OSError, ValueError, PermissionError):
                continue
        if self.pid is None:
            app = Application(backend="uia").start(str(self.args.exe), work_dir=str(self.args.exe.parent))
            self.pid = app.process
        if not self.largest_window():
            deadline = time.monotonic() + self.args.wait_seconds
            while time.monotonic() < deadline and not self.largest_window():
                time.sleep(.4)
        if not self.largest_window():
            raise RuntimeError("ไม่พบหน้าต่าง HRMI หลังเปิดโปรแกรม")
        self.done(f"ProcessId: {self.pid}")

    def login(self, user: str, password: str) -> None:
        self.start("เข้าสู่ระบบ HRMI")
        dialog = self.window("Login")
        if not dialog:
            raise RuntimeError("ไม่พบหน้าต่าง Login")
        self.activate(dialog)
        # HRMI starts focus on Login button; same tab order as the existing workflow.
        keyboard.send_keys("+{TAB}+{TAB}^a")
        keyboard.send_keys(literal_keys(password), with_spaces=True, pause=.02, vk_packet=True)
        keyboard.send_keys("+{TAB}^a")
        keyboard.send_keys(literal_keys(user), with_spaces=True, pause=.02, vk_packet=True)
        keyboard.send_keys("{TAB}{TAB}{ENTER}")
        deadline = time.monotonic() + max(10, self.args.wait_seconds)
        while time.monotonic() < deadline:
            if self.window("Organization"):
                self.done("พบหน้า Organization แล้ว")
                return
            alert = self.window("Alert") or self.window("Information")
            if alert:
                raise RuntimeError(f"HRMI แสดงข้อความหลัง Login: {alert.window_text()}")
            time.sleep(.5)
        raise RuntimeError("ไม่พบหน้าเลือกองค์กรหลัง Login")

    def select_org(self, code: str) -> None:
        self.org = code
        self.start("เลือกองค์กร", code)
        dialog = self.wait_window("Organization", self.args.wait_seconds)
        if not dialog:
            raise RuntimeError("ไม่พบหน้าต่าง Organization")
        self.activate(dialog)
        rect = dialog.rectangle()
        if rect.width() < 250 or rect.height() < 180:
            raise RuntimeError("หน้าต่าง Organization มีขนาดผิดปกติ")
        # Custom grid does not expose row selection through UIA.
        row = ORGANIZATIONS[code]
        self.click_owned(int(rect.left + rect.width() * .32),
                         int(rect.top + rect.height() * (.245 + row * .072)))
        time.sleep(.2)
        buttons = [e for e in dialog.descendants() if e.window_text().strip().casefold() == "ok"]
        if buttons:
            try:
                buttons[0].click()
            except Exception:
                pass
        time.sleep(.35)
        if self.window("Organization"):
            self.click_owned(int(rect.left + rect.width() * .72),
                             int(rect.top + rect.height() * .91))
        if self.wait_window("Organization", 2):
            raise RuntimeError(f"หน้าต่าง Organization ไม่ปิดหลังเลือก {code}")
        self.done("ยืนยันแถวองค์กรแล้ว")

    def navigate(self) -> None:
        self.start("เปิดเมนู Time Attendance")
        alert = self.window("Alert")
        if alert:
            alert.close()
            if self.wait_window("Alert", 1):
                raise RuntimeError("ปิดหน้าต่าง Alert ไม่สำเร็จ")
        named = self.elements("Time Attendance", exact=True)
        for item in named:
            try:
                item.select() if item.element_info.control_type in ("ListItem", "TreeItem") else item.invoke()
                if self.wait_text("Time Attendance Flow", 3):
                    self.done("UI Automation")
                    return
            except Exception:
                continue
        if not self.args.calibrate_time_attendance:
            raise RuntimeError("เมนู Time Attendance เป็น custom control; ใช้ --calibrate-time-attendance")
        nav = next((e for e in self.elements("Navigator Explorer Bar") if e.is_visible()), None)
        if not nav:
            raise RuntimeError("ไม่พบ Navigator Explorer Bar")
        bounds = nav.rectangle()
        config = None
        if CONFIG.exists():
            try:
                config = json.loads(CONFIG.read_text(encoding="utf-8-sig"))
            except (ValueError, OSError):
                pass
        if config and config.get("Version") == 3 and config.get("NavigatorRelative"):
            try:
                rx = float(config["NavigatorRelative"]["X"])
                ry = float(config["NavigatorRelative"]["Y"])
                if not (0 <= rx <= 1 and 0 <= ry <= 1):
                    raise ValueError("พิกัดสัมพัทธ์ไม่ถูกต้อง")
                self.activate(self.largest_window())
                self.click_owned(int(bounds.left + bounds.width() * rx),
                                 int(bounds.top + bounds.height() * ry))
                if self.wait_text("Time Attendance Flow", 3):
                    self.done("พิกัดสัมพัทธ์ Navigator")
                    return
            except (ValueError, RuntimeError) as exc:
                self.log(self.step, "INFO", f"ปรับเทียบใหม่: {exc}")
        print("คลิกเมนู Time Attendance ในหน้าต่าง HRMI ด้วยตนเองหนึ่งครั้ง (รอ 120 วินาที)")
        deadline = time.monotonic() + 120
        was_down = bool(win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000)
        point = None
        while time.monotonic() < deadline:
            is_down = bool(win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000)
            if is_down and not was_down:
                candidate = win32api.GetCursorPos()
                handle = win32gui.WindowFromPoint(candidate)
                _, pid = win32process.GetWindowThreadProcessId(handle)
                if pid == self.pid:
                    point = candidate
                    break
            was_down = is_down
            time.sleep(.02)
        if not point or not self.wait_text("Time Attendance Flow", 5):
            raise RuntimeError("ปรับเทียบไม่สำเร็จ: ไม่พบการคลิกหรือหน้า Time Attendance Flow")
        rx = (point[0] - bounds.left) / bounds.width()
        ry = (point[1] - bounds.top) / bounds.height()
        if not (0 <= rx <= 1 and 0 <= ry <= 1):
            raise RuntimeError("จุดที่คลิกไม่อยู่ใน Navigator")
        CONFIG.write_text(json.dumps({"Version": 3, "NavigatorRelative": {"X": rx, "Y": ry},
                                      "VerifiedSubmenu": "Time Attendance Flow",
                                      "CalibratedAt": datetime.now().astimezone().isoformat()},
                                     ensure_ascii=False, indent=2), encoding="utf-8")
        self.done("ผู้ใช้คลิกและบันทึกพิกัด Navigator แล้ว")

    def open_processing(self) -> None:
        self.navigate()
        self.start("เปิดหน้าประมวลผลเวลาการทำงาน")
        self.activate(self.largest_window())
        items = self.elements("ประมวลผลเวลาการทำงาน")
        if not items:
            raise RuntimeError("ไม่พบเมนูประมวลผลเวลาการทำงาน")
        item = items[0]
        try:
            item.select() if item.element_info.control_type in ("ListItem", "TreeItem") else item.invoke()
        except Exception:
            # Focusing HRMI may move its window. Read the rectangle after focus.
            rect = item.rectangle()
            if rect.width() <= 4 or rect.height() <= 4:
                raise RuntimeError("เมนูประมวลผลไม่มีพื้นที่คลิก")
            self.click_owned(int(rect.mid_point().x), int(rect.mid_point().y))
        deadline = time.monotonic() + 8
        processing_open = False
        while time.monotonic() < deadline:
            processing_open = any(e.element_info.control_type == "Window"
                                  for e in self.elements("ประมวลผลเวลาการทำงาน"))
            if processing_open:
                break
            time.sleep(.3)
        if not processing_open:
            raise RuntimeError("ไม่พบหน้าต่างประมวลผลเวลาการทำงาน")
        self.done("เปิดหน้าประมวลผลแล้ว; ยังไม่ได้กดเริ่มคำนวณ")
        self.set_processing_dates()

    def control_by_id(self, automation_id: str):
        return next((e for w in self.windows() for e in w.descendants()
                     if e.element_info.automation_id == automation_id and e.is_visible()), None)

    def popup_below(self, control_rect):
        for popup in self.windows():
            rect = popup.rectangle()
            if (not popup.window_text() and abs(rect.left - control_rect.left) <= 5
                    and abs(rect.top - control_rect.bottom) <= 5
                    and 180 <= rect.width() <= 280 and 175 <= rect.height() <= 260):
                return popup
        raise RuntimeError("ไม่พบปฏิทินหรือรายการเดือนของ HRMI")

    def choose_calendar_date(self, automation_id: str) -> None:
        control = self.control_by_id(automation_id)
        if not control:
            raise RuntimeError(f"ไม่พบช่องวันที่ {automation_id}")
        source = parse_thai_date(control.window_text()) or datetime.now(timezone(timedelta(hours=7))).date()
        rect = control.rectangle()
        self.click_owned(rect.right - 10, rect.top + 10)
        time.sleep(.2)
        popup = self.popup_below(rect)
        box = popup.rectangle()
        month_delta = (self.target_date.year - source.year) * 12 + self.target_date.month - source.month
        if abs(month_delta) > 2:
            raise RuntimeError(f"เดือนในปฏิทินต่างจากวันที่เป้าหมายเกิน 2 เดือน: {source}")
        for _ in range(abs(month_delta)):
            arrow_x = box.right - 15 if month_delta > 0 else box.left + 15
            self.click_owned(arrow_x, box.top + 33)
            time.sleep(.15)
        first_weekday, _ = calendar.monthrange(self.target_date.year, self.target_date.month)
        cell = first_weekday + self.target_date.day - 1
        column, row = cell % 7, cell // 7
        x = box.left + round(box.width() * (37 + 25.3 * column) / 220)
        y = box.top + round(box.height() * (83 + 17.3 * row) / 211)
        self.click_owned(x, y)
        time.sleep(.25)
        actual = self.control_by_id(automation_id)
        if not actual or parse_thai_date(actual.window_text()) != self.target_date:
            shown = actual.window_text() if actual else "ไม่พบช่อง"
            raise RuntimeError(f"ตั้งวันที่ {automation_id} ไม่สำเร็จ: {shown}")

    def choose_month(self) -> None:
        control = self.control_by_id("devLookUpMonth")
        if not control:
            raise RuntimeError("ไม่พบช่องเดือน")
        wanted = THAI_MONTHS[self.target_date.month - 1]
        if control.window_text().strip() == wanted:
            return
        rect = control.rectangle()
        self.click_owned(rect.right - 10, rect.top + 10)
        time.sleep(.2)
        popup = self.popup_below(rect).rectangle()
        if not 180 <= popup.height() <= 210:
            raise RuntimeError("ขนาดรายการเดือนผิดปกติ")
        self.click_owned(popup.left + 45, popup.top + 12 + (self.target_date.month - 1) * 15)
        time.sleep(.25)
        actual = self.control_by_id("devLookUpMonth")
        if not actual or actual.window_text().strip() != wanted:
            raise RuntimeError(f"ตั้งเดือน {wanted} ไม่สำเร็จ")

    def set_processing_dates(self) -> None:
        self.start("ตั้งวันที่ประมวลผล", thai_date_text(self.target_date))
        self.choose_calendar_date("devDtStartDate")
        self.choose_calendar_date("devDtEndDate")
        self.choose_month()
        self.done(f"วันเริ่มและวันสิ้นสุด: {thai_date_text(self.target_date)}; เดือน: {THAI_MONTHS[self.target_date.month - 1]}")

    def wizard_control(self, automation_id: str):
        wizard = self.control_by_id("frmTimeRecordWiz")
        if not wizard:
            raise RuntimeError("ไม่พบหน้าต่างประมวลผลเวลา")
        return next((e for e in wizard.descendants()
                     if e.element_info.automation_id == automation_id), None)

    def wizard_step(self) -> str:
        item = self.wizard_control("lblStepName")
        return item.window_text().strip() if item else ""

    def press_wizard_next(self, expected_step: str | None = None) -> None:
        button = self.wizard_control("devBtnNext")
        if not button or not button.is_enabled():
            raise RuntimeError("ปุ่ม Next ใช้งานไม่ได้")
        button.click_input()
        if expected_step:
            deadline = time.monotonic() + self.args.wait_seconds
            while time.monotonic() < deadline:
                if self.wizard_step() == expected_step:
                    return
                time.sleep(.3)
            raise RuntimeError(f"กด Next แล้วไม่พบขั้น {expected_step}; พบ {self.wizard_step()}")

    def advance_wizard(self) -> None:
        self.start("กด Next ในขั้นตอนประมวลผล", self.org)
        if not self.wizard_control("devDtStartDate") or not self.wizard_control("devDtEndDate"):
            raise RuntimeError(f"ไม่พบช่องวันที่ในขั้นเริ่มต้น: {self.wizard_step()}")
        self.press_wizard_next("เลือกช่วงข้อมูลพนักงาน")
        self.log(self.step, "INFO", "ผ่านหน้าเลือกช่วงข้อมูลพนักงาน")
        self.press_wizard_next("สรุปก่อนประมวลผล")
        summary_date = self.wizard_control("devlblStartDate")
        expected = f"{self.target_date:%d/%m}/{self.target_date.year + 543}  -  {self.target_date:%d/%m}/{self.target_date.year + 543}"
        if not summary_date or summary_date.window_text().strip() != expected:
            shown = summary_date.window_text() if summary_date else "ไม่พบวันที่"
            raise RuntimeError(f"ช่วงวันที่ในหน้าสรุปไม่ตรง: {shown}")
        self.log(self.step, "INFO", f"หน้าสรุปแสดงวันที่ {expected}; เริ่มประมวลผล")
        self.press_wizard_next()
        deadline = time.monotonic() + 20 * 60
        while time.monotonic() < deadline:
            question = self.control_by_id("frmMessageBox")
            if question or self.wizard_step() == "สรุปผลหลังประมวลผล":
                break
            time.sleep(3)
        else:
            raise RuntimeError("รอการประมวลผลเกิน 20 นาที")
        question = self.control_by_id("frmMessageBox")
        if question:
            message = next((e.window_text() for e in question.descendants()
                            if e.element_info.automation_id == "lblErrorMessage"), "")
            if "เสร็จสิ้นแล้ว" not in message or "ต้องการตรวจสอบ" not in message:
                raise RuntimeError(f"HRMI แสดงคำถามที่ไม่รู้จัก: {message}")
            no_button = next((e for e in question.descendants()
                              if e.element_info.automation_id == "btnNo"), None)
            if not no_button:
                raise RuntimeError("ไม่พบปุ่ม No ในคำถามหลังประมวลผล")
            no_button.click_input()
            self.log(self.step, "INFO", "ประมวลผลเสร็จ; เลือก No สำหรับการเปิดหน้าตรวจสอบเพิ่มเติม")
        deadline = time.monotonic() + self.args.wait_seconds
        while time.monotonic() < deadline and self.wizard_step() != "สรุปผลหลังประมวลผล":
            time.sleep(.3)
        if self.wizard_step() != "สรุปผลหลังประมวลผล":
            raise RuntimeError(f"ไม่พบหน้าสรุปหลังประมวลผล: {self.wizard_step()}")
        finish = self.wizard_control("devBtnFinish")
        if not finish or not finish.is_enabled():
            raise RuntimeError("ไม่พบปุ่ม Finish ที่ใช้งานได้")
        finish.click_input()
        deadline = time.monotonic() + self.args.wait_seconds
        while time.monotonic() < deadline:
            if not self.control_by_id("frmTimeRecordWiz"):
                self.done("กด Finish และหน้าต่างประมวลผลปิดแล้ว")
                return
            time.sleep(.3)
        raise RuntimeError("กด Finish แล้ว แต่หน้าต่างประมวลผลยังไม่ปิด")

    def return_to_picker(self, user: str, password: str) -> None:
        self.start("กลับไปหน้าเลือกองค์กร")
        processing = next((e for e in self.elements("ประมวลผลเวลาการทำงาน", exact=True)
                           if e.element_info.control_type == "Window"), None)
        if processing:
            processing.close()
            time.sleep(.6)
        alert = self.window("Alert")
        if alert:
            self.log(self.step, "INFO", "ปิด Alert ที่ HRMI แสดงหลังปิดหน้าประมวลผล")
            alert.close()
            time.sleep(.4)
        main = self.largest_window()
        self.activate(main)
        file_menu = next((e for e in main.descendants()
                          if e.window_text() == "File" and e.element_info.control_type == "MenuItem"), None)
        if not file_menu:
            raise RuntimeError("ไม่พบเมนู File")
        file_menu.click_input()
        time.sleep(.3)
        menu_rect = file_menu.rectangle()
        # HRMI's legacy popup does not expose its Login entry through UIA.
        self.click_owned(menu_rect.left + 48, menu_rect.bottom + 15)
        time.sleep(.5)
        alert = self.window("Alert")
        if alert:
            self.log(self.step, "INFO", "HRMI แสดง Alert ระหว่างกลับหน้า Login; ปิดและลองใหม่")
            alert.close()
            time.sleep(.4)
            file_menu.click_input()
            time.sleep(.3)
            self.click_owned(menu_rect.left + 48, menu_rect.bottom + 15)
        deadline = time.monotonic() + self.args.wait_seconds
        while time.monotonic() < deadline:
            if self.window("Organization"):
                self.done("พบหน้า Organization")
                return
            if self.window("Login"):
                if not self.args.auto_login:
                    raise RuntimeError("HRMI ขอ Login อีกครั้ง; ใช้ --auto-login")
                self.login(user, password)
                if not self.window("Organization"):
                    raise RuntimeError("ไม่พบหน้า Organization หลัง Login")
                return
            time.sleep(.3)
        raise RuntimeError("ไม่พบหน้าเลือกองค์กรหลัง File > Login")


def main() -> int:
    parser = argparse.ArgumentParser(description="Run HRMI time processing wizard for all configured organizations")
    parser.add_argument("--exe", type=Path, default=DEFAULT_EXE)
    parser.add_argument("--wait-seconds", type=int, default=30)
    parser.add_argument("--organizations", nargs="+", default=list(ORGANIZATIONS), choices=list(ORGANIZATIONS))
    parser.add_argument("--log-path", type=Path, default=ROOT / "logs")
    parser.add_argument("--calibrate-time-attendance", action="store_true")
    parser.add_argument("--auto-login", action="store_true")
    parser.add_argument("--stop-before-next", action="store_true",
                        help="open and set dates only; do not advance the wizard")
    args = parser.parse_args()
    load_env()
    user = os.getenv("HRMI_USER") or os.getenv("user", "")
    password = os.getenv("HRMI_PASSWORD") or os.getenv("password", "")
    if args.auto_login and (not user or not password):
        parser.error("ตั้ง HRMI_USER และ HRMI_PASSWORD ก่อนใช้ --auto-login")
    if args.auto_login and (user.strip().casefold() in {"ชื่อผู้ใช้", "username", "user name", "ใส่ชื่อผู้ใช้"}
                            or password.strip().casefold() in {"รหัสผ่าน", "password", "ใส่รหัสผ่าน"}):
        parser.error("ข้อมูล Login ยังเป็นข้อความตัวอย่าง")
    hrmi = Hrmi(args)
    try:
        hrmi.log("เริ่มต้น", "INFO", f"Log: {hrmi.log_path}")
        hrmi.connect()
        if hrmi.window("Login"):
            if not args.auto_login:
                raise RuntimeError("HRMI อยู่หน้า Login; ใช้ --auto-login หรือ Login เองก่อนรัน")
            hrmi.login(user, password)
        if not hrmi.window("Organization"):
            raise RuntimeError("ไม่พบหน้าเลือกองค์กร; ตรวจยืนยันองค์กรปัจจุบันไม่ได้")
        for index, org in enumerate(args.organizations):
            org_started = datetime.now(BANGKOK_TZ)
            try:
                hrmi.select_org(org)
                hrmi.open_processing()
                if not args.stop_before_next:
                    hrmi.advance_wizard()
                message = ("ประมวลผลและกด Finish แล้ว" if not args.stop_before_next
                           else "เปิดหน้าประมวลผลและตั้งวันที่แล้ว")
                hrmi.log("ประมวลผลองค์กร", "OK", message)
                hrmi.record_timing(org, org_started, datetime.now(BANGKOK_TZ), "สำเร็จ", message)
            except Exception as exc:
                hrmi.record_timing(org, org_started, datetime.now(BANGKOK_TZ), "ผิดพลาด", str(exc))
                raise
            if index < len(args.organizations) - 1:
                hrmi.return_to_picker(user, password)
        hrmi.log("จบสคริปต์", "INFO", f"เสร็จ {len(args.organizations)} องค์กร")
        return 0
    except Exception as exc:
        hrmi.log(hrmi.step, "FATAL", str(exc))
        print(f"หยุด: {exc}", file=sys.stderr)
        return 1
    finally:
        hrmi.close()


if __name__ == "__main__":
    raise SystemExit(main())
