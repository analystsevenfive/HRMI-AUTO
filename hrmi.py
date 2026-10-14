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
from contextlib import AbstractContextManager
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from pywinauto import Application, Desktop, keyboard, mouse
import win32api
import win32con
import win32event
import win32gui
import win32process
import winerror


ROOT = Path(__file__).resolve().parent
DEFAULT_EXE = Path(r"C:\Program Files (x86)\Prosoft\Prosoft HRMI\Bin\HRMI.exe")
ORGANIZATIONS = {"[BB]": 0, "[BK]": 1, "[NTS]": 2, "[PK]": 3, "[SM]": 4}
CONFIG = ROOT / "hrmi-navigation.json"
THAI_MONTHS = ("มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม")
BANGKOK_TZ = timezone(timedelta(hours=7))
COMPLETED_MESSAGE = "ประมวลผลและกด Finish แล้ว"
MUTEX_NAME = r"Local\HRMI_AUTO_DAILY"
ABSENCE_REPORT_TYPES = {"ขาด": 0, "สาย": 1}
REPORT_TYPES = tuple(ABSENCE_REPORT_TYPES)


def bangkok_target_date() -> date:
    return datetime.now(BANGKOK_TZ).date() - timedelta(days=2)


def scheduled_processing_range(run_date: date | None = None) -> tuple[date, date]:
    """Return the unattended processing range for a Bangkok calendar day."""
    run_date = run_date or datetime.now(BANGKOK_TZ).date()
    if run_date.weekday() == calendar.TUESDAY:
        return run_date - timedelta(days=3), run_date - timedelta(days=2)
    target = run_date - timedelta(days=2)
    return target, target


def parse_cli_date(value: str) -> date:
    value = value.strip()
    try:
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
            return date.fromisoformat(value)
        match = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", value)
        if match:
            year = int(match.group(3))
            if year >= 2400:
                year -= 543
            return date(year, int(match.group(2)), int(match.group(1)))
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"วันที่ไม่ถูกต้อง: {value}") from exc
    raise argparse.ArgumentTypeError(
        f"รูปแบบวันที่ไม่ถูกต้อง: {value} (ใช้ YYYY-MM-DD หรือ DD/MM/YYYY)"
    )


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


def load_timing_rows(path: Path) -> list[dict]:
    if not path.exists():
        return []
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    return value if isinstance(value, list) else []


def completed_organizations(path: Path, start_date: date, end_date: date | None = None) -> set[str]:
    end_date = end_date or start_date
    wanted_start = start_date.isoformat()
    wanted_end = end_date.isoformat()
    def same_range(row: dict) -> bool:
        row_start = row.get("target_start_date") or row.get("target_date")
        row_end = row.get("target_end_date") or row.get("target_date")
        return row_start == wanted_start and row_end == wanted_end
    return {
        str(row.get("organization"))
        for row in load_timing_rows(path)
        if same_range(row)
        and row.get("status") == "สำเร็จ"
        and "Finish" in str(row.get("message", ""))
    }


def organizations_to_run(requested: list[str], completed: set[str], *,
                         scheduled: bool, force: bool,
                         input_fn=input) -> list[str]:
    if force:
        return list(requested)
    pending = [org for org in requested if org not in completed]
    if pending or scheduled:
        return pending
    answer = input_fn("วันที่เป้าหมายนี้ประมวลผลครบแล้ว ต้องการรันซ้ำหรือไม่? [y/N]: ")
    return list(requested) if answer.strip().casefold() in {"y", "yes"} else []


class SingleInstance(AbstractContextManager):
    def __init__(self, name: str = MUTEX_NAME):
        self.handle = win32event.CreateMutex(None, True, name)
        self.already_running = win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if not self.already_running:
            win32event.ReleaseMutex(self.handle)
        win32api.CloseHandle(self.handle)


class Hrmi:
    def __init__(self, args: argparse.Namespace):
        self.args = args
        self.pid: int | None = None
        self.org = ""
        self.step = "เริ่มต้น"
        # Every automatic flow uses the same date rule. Explicit custom dates
        # supplied on the command line still take precedence below.
        default_start, default_end = scheduled_processing_range()
        self.start_date = args.start_date or default_start
        self.end_date = args.end_date or default_end
        self.target_date = self.start_date
        self.run_id = datetime.now(BANGKOK_TZ).strftime("%Y%m%d-%H%M%S")
        self.timing_json = args.log_path / "HRMI-processing-times.json"
        self.timing_xlsx = ROOT / "HRMI-processing-times.xlsx"
        self.report_log = args.log_path / "HRMI-report-exports.json"
        self.report_timing_xlsx = ROOT / "HRMI-report-export-times.xlsx"
        self.export_path = getattr(args, "export_path", ROOT / "exports")
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
        rows = load_timing_rows(self.timing_json)
        rows.append({
            "run_id": self.run_id,
            "organization": organization,
            "target_date": self.target_date.isoformat(),
            "target_start_date": self.start_date.isoformat(),
            "target_end_date": self.end_date.isoformat(),
            "started_at": started_at.isoformat(),
            "finished_at": finished_at.isoformat(),
            "duration_seconds": round((finished_at - started_at).total_seconds(), 3),
            "status": status,
            "message": message,
        })
        self.timing_json.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        builder = ROOT / "build_timing_excel.py"
        try:
            subprocess.run(
                [sys.executable, str(builder), str(self.timing_json), str(self.timing_xlsx), "processing"],
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
            # Focus can be stolen after activation, for example by an editor.
            # Restore only an HRMI window containing the intended point, then
            # repeat the ownership check before sending any mouse input.
            candidates = []
            for window in self.windows():
                rect = window.rectangle()
                if rect.left <= x < rect.right and rect.top <= y < rect.bottom:
                    candidates.append(window)
            if candidates:
                owner = min(candidates, key=lambda w: w.rectangle().width() * w.rectangle().height())
                bounds = owner.rectangle()
                self.activate(owner)
                current = owner.rectangle()
                if (current.left, current.top, current.right, current.bottom) != (bounds.left, bounds.top, bounds.right, bounds.bottom):
                    raise RuntimeError("หน้าต่าง HRMI ย้ายตำแหน่งระหว่างเตรียมคลิก; กรุณารันใหม่")
                handle = win32gui.WindowFromPoint((x, y))
                _, pid = win32process.GetWindowThreadProcessId(handle)
            if pid != self.pid:
                raise RuntimeError(f"ตำแหน่งคลิก ({x},{y}) ไม่ได้อยู่บนหน้าต่าง HRMI; มีหน้าต่างอื่นบังอยู่")
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
        deadline = time.monotonic() + self.args.wait_seconds
        main = self.largest_window()
        while main is None and time.monotonic() < deadline:
            time.sleep(.3)
            main = self.largest_window()
        if main is None:
            raise RuntimeError("ไม่พบหน้าหลัก HRMI หลังเลือกองค์กร")
        try:
            main.maximize()
            time.sleep(.3)
        except Exception:
            pass
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
        if config and config.get("Version") == 2 and config.get("TimeAttendance"):
            try:
                point = config["TimeAttendance"]
                self.activate(self.largest_window())
                self.click_owned(int(point["X"]), int(point["Y"]))
                if self.wait_text("Time Attendance Flow", 3):
                    rx = (int(point["X"]) - bounds.left) / bounds.width()
                    ry = (int(point["Y"]) - bounds.top) / bounds.height()
                    if 0 <= rx <= 1 and 0 <= ry <= 1:
                        CONFIG.write_text(json.dumps({
                            "Version": 3,
                            "NavigatorRelative": {"X": rx, "Y": ry},
                            "VerifiedSubmenu": "Time Attendance Flow",
                            "CalibratedAt": datetime.now().astimezone().isoformat(),
                        }, ensure_ascii=False, indent=2), encoding="utf-8")
                    self.done("พิกัดเดิมและแปลงเป็นพิกัดสัมพัทธ์ Navigator")
                    return
            except (KeyError, TypeError, ValueError, RuntimeError) as exc:
                self.log(self.step, "INFO", f"ใช้พิกัดเดิมไม่สำเร็จ: {exc}")
        if not self.args.calibrate_time_attendance:
            raise RuntimeError("เมนู Time Attendance เป็น custom control; ใช้ --calibrate-time-attendance")
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

    def choose_combo_index(self, control, index: int) -> None:
        rect = control.rectangle()
        self.click_owned(rect.right - 10, int(rect.mid_point().y))
        time.sleep(.3)
        items = sorted(
            (e for w in self.windows() for e in w.descendants(control_type="ListItem")
             if e.is_visible()),
            key=lambda e: e.rectangle().top,
        )
        if index >= len(items):
            raise RuntimeError(f"ตัวเลือกในรายการไม่ครบ: ต้องการลำดับ {index + 1}, พบ {len(items)}")
        items[index].click_input()
        time.sleep(.3)

    def choose_combo_text(self, control, wanted: str) -> None:
        normalize = lambda value: "".join(value.split()).casefold()
        if normalize(control.window_text()) == normalize(wanted):
            return
        rect = control.rectangle()
        item = None
        for _ in range(3):
            keyboard.send_keys("{ESC}")
            self.click_owned(rect.right - 10, int(rect.mid_point().y))
            deadline = time.monotonic() + 2
            while time.monotonic() < deadline:
                item = next(
                    (e for w in self.windows() for e in w.descendants(control_type="ListItem")
                     if e.is_visible() and normalize(e.window_text()) == normalize(wanted)),
                    None,
                )
                if item:
                    break
                time.sleep(.2)
            if item:
                break
        if not item:
            raise RuntimeError(f"ไม่พบตัวเลือก {wanted}")
        item.click_input()
        time.sleep(.3)
        if normalize(control.window_text()) != normalize(wanted):
            raise RuntimeError(f"ตั้งตัวเลือก {wanted} ไม่สำเร็จ: {control.window_text()}")

    def popup_below(self, control_rect):
        deadline = time.monotonic() + 3
        while time.monotonic() < deadline:
            for window in self.windows():
                for popup in [window, *window.descendants()]:
                    rect = popup.rectangle()
                    if (popup.is_visible() and abs(rect.left - control_rect.left) <= 8
                            and abs(rect.top - control_rect.bottom) <= 8
                            and 180 <= rect.width() <= 280 and 175 <= rect.height() <= 260):
                        return popup
            handles = []
            def collect(handle, _):
                if (win32gui.IsWindowVisible(handle)
                        and win32process.GetWindowThreadProcessId(handle)[1] == self.pid):
                    handles.append(handle)
            win32gui.EnumWindows(collect, None)
            for handle in list(handles):
                win32gui.EnumChildWindows(handle, collect, None)
            for handle in dict.fromkeys(handles):
                left, top, right, bottom = win32gui.GetWindowRect(handle)
                if (abs(left - control_rect.left) <= 8
                        and abs(top - control_rect.bottom) <= 8
                        and 180 <= right - left <= 280 and 175 <= bottom - top <= 260):
                    return Desktop(backend="win32").window(handle=handle).wrapper_object()
            time.sleep(.15)
        raise RuntimeError("ไม่พบปฏิทินหรือรายการเดือนของ HRMI")

    def choose_calendar_date(self, automation_id: str, target_date: date) -> None:
        control = self.control_by_id(automation_id)
        if not control:
            raise RuntimeError(f"ไม่พบช่องวันที่ {automation_id}")
        source = parse_thai_date(control.window_text()) or datetime.now(timezone(timedelta(hours=7))).date()
        rect = control.rectangle()
        self.click_owned(rect.right - 10, rect.top + 10)
        time.sleep(.2)
        popup = self.popup_below(rect)
        box = popup.rectangle()
        month_delta = (target_date.year - source.year) * 12 + target_date.month - source.month
        if abs(month_delta) > 2:
            raise RuntimeError(f"เดือนในปฏิทินต่างจากวันที่เป้าหมายเกิน 2 เดือน: {source}")
        for _ in range(abs(month_delta)):
            arrow_x = box.right - 15 if month_delta > 0 else box.left + 15
            self.click_owned(arrow_x, box.top + 33)
            time.sleep(.15)
        first_weekday, _ = calendar.monthrange(target_date.year, target_date.month)
        cell = first_weekday + target_date.day - 1
        column, row = cell % 7, cell // 7
        x = box.left + round(box.width() * (37 + 25.3 * column) / 220)
        y = box.top + round(box.height() * (83 + 17.3 * row) / 211)
        self.click_owned(x, y)
        time.sleep(.25)
        actual = self.control_by_id(automation_id)
        if not actual or parse_thai_date(actual.window_text()) != target_date:
            shown = actual.window_text() if actual else "ไม่พบช่อง"
            raise RuntimeError(f"ตั้งวันที่ {automation_id} ไม่สำเร็จ: {shown}")

    def choose_month(self) -> None:
        control = self.control_by_id("devLookUpMonth")
        if not control:
            raise RuntimeError("ไม่พบช่องเดือน")
        wanted = THAI_MONTHS[self.end_date.month - 1]
        if control.window_text().strip() == wanted:
            return
        rect = control.rectangle()
        self.click_owned(rect.right - 10, rect.top + 10)
        time.sleep(.2)
        popup = self.popup_below(rect).rectangle()
        if not 180 <= popup.height() <= 210:
            raise RuntimeError("ขนาดรายการเดือนผิดปกติ")
        self.click_owned(popup.left + 45, popup.top + 12 + (self.start_date.month - 1) * 15)
        time.sleep(.25)
        actual = self.control_by_id("devLookUpMonth")
        if not actual or actual.window_text().strip() != wanted:
            raise RuntimeError(f"ตั้งเดือน {wanted} ไม่สำเร็จ")

    def set_processing_dates(self) -> None:
        self.start("ตั้งวันที่ประมวลผล", f"{thai_date_text(self.start_date)} - {thai_date_text(self.end_date)}")
        self.choose_calendar_date("devDtStartDate", self.start_date)
        self.choose_calendar_date("devDtEndDate", self.end_date)
        self.choose_month()
        self.done(f"ช่วงวันที่: {thai_date_text(self.start_date)} - {thai_date_text(self.end_date)}; เดือน: {THAI_MONTHS[self.end_date.month - 1]}")

    def report_file(self, organization: str, report_type: str) -> Path:
        org = organization.strip("[]")
        date_part = self.start_date.isoformat()
        if self.end_date != self.start_date:
            date_part += f"_{self.end_date.isoformat()}"
        return self.export_path / date_part / f"HRMI_{org}_{report_type}_{date_part}.xlsx"

    def report_is_complete(self, organization: str, report_type: str) -> bool:
        path = self.report_file(organization, report_type)
        return path.is_file() and path.stat().st_size > 0

    def record_report_export(self, organization: str, report_type: str,
                             path: Path, status: str, details: str = "",
                             started_at: datetime | None = None,
                             finished_at: datetime | None = None) -> None:
        finished_at = finished_at or datetime.now(BANGKOK_TZ)
        started_at = started_at or finished_at
        rows = load_timing_rows(self.report_log)
        rows.append({
            "run_id": self.run_id,
            "organization": organization,
            "report_type": report_type,
            "target_start_date": self.start_date.isoformat(),
            "target_end_date": self.end_date.isoformat(),
            "started_at": started_at.isoformat(),
            "finished_at": finished_at.isoformat(),
            "duration_seconds": round((finished_at - started_at).total_seconds(), 3),
            "exported_at": finished_at.isoformat(),
            "path": str(path),
            "status": status,
            "details": details,
        })
        self.report_log.parent.mkdir(parents=True, exist_ok=True)
        self.report_log.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        builder = ROOT / "build_timing_excel.py"
        try:
            subprocess.run(
                [sys.executable, str(builder), str(self.report_log), str(self.report_timing_xlsx), "reports"],
                cwd=ROOT, check=True, text=True,
                stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
            )
        except (OSError, subprocess.CalledProcessError) as exc:
            details = exc.stderr.strip() if isinstance(exc, subprocess.CalledProcessError) and exc.stderr else str(exc)
            self.log("สร้าง Excel เวลา Export", "WARN", details)

    def report_preview(self):
        return next((w for w in self.windows()
                     if w.element_info.automation_id == "frmPreview"), None)

    def open_absence_report(self) -> None:
        alert = next((w for w in self.windows()
                      if w.element_info.automation_id == "frmAlert"), None)
        if alert:
            alert.close()
            time.sleep(.4)
        preview = self.report_preview()
        if preview:
            if any(e.element_info.automation_id == "frmReportRangeForm"
                   for e in preview.descendants()):
                return
            range_button = next((e for e in preview.descendants(control_type="Button")
                                 if e.window_text() == "Range"), None)
            if range_button:
                range_button.click_input()
        else:
            main = self.largest_window()
            try:
                main.maximize()
            except Exception:
                pass
            self.activate(main)
            tree = next((e for e in main.descendants(control_type="Tree") if e.is_visible()), None)
            if not tree:
                raise RuntimeError("ไม่พบรายการเมนู Time Attendance")
            point = tree.rectangle().mid_point()
            mouse.scroll(coords=(int(point.x), int(point.y)), wheel_dist=100)
            time.sleep(.5)
            line_down = next((e for e in tree.descendants(control_type="Button")
                              if e.window_text() == "Line Down"), None)
            if not line_down:
                raise RuntimeError("ไม่พบปุ่มเลื่อนรายการรายงาน")
            for _ in range(10):
                line_down.click_input()
                time.sleep(.08)
            main = self.largest_window()
            item = next((e for e in main.descendants(control_type="TreeItem")
                         if e.window_text() == "Node26" and e.is_visible()), None)
            if not item:
                raise RuntimeError("ไม่พบเมนูรายงานการลา ขาด สาย (Node26)")
            tree_rect = tree.rectangle()
            item_rect = item.rectangle()
            if item_rect.top < tree_rect.top or item_rect.bottom > tree_rect.bottom:
                raise RuntimeError("เมนูรายงานการลา ขาด สายอยู่นอกพื้นที่คลิก")
            item.double_click_input()
        deadline = time.monotonic() + max(self.args.wait_seconds, 90)
        while time.monotonic() < deadline:
            preview = self.report_preview()
            if preview and any(e.element_info.automation_id == "frmReportRangeForm"
                               for e in preview.descendants()):
                return
            time.sleep(.3)
        raise RuntimeError("ไม่พบหน้ากำหนดช่วงรายงานการลา ขาด สาย")

    def export_absence_report(self, organization: str, report_type: str) -> Path:
        output = self.report_file(organization, report_type)
        output.parent.mkdir(parents=True, exist_ok=True)
        started_at = datetime.now(BANGKOK_TZ)
        self.start("Export รายงาน", f"{organization} {report_type} -> {output}")
        self.open_absence_report()
        preview = self.report_preview()
        if not preview:
            raise RuntimeError("ไม่พบหน้ารายงาน")
        controls = {e.element_info.automation_id: e for e in preview.descendants()
                    if e.element_info.automation_id}
        report_combo = controls.get("cmbReportList")
        data_combo = controls.get("Combo05")
        if not report_combo or not data_combo:
            raise RuntimeError("ไม่พบตัวเลือกประเภทรายงานหรือประเภทข้อมูล")
        self.choose_combo_text(report_combo, "แบบแจกแจง(นับสายเช้า-บ่าย)")
        self.choose_combo_text(data_combo, report_type)
        selected_data = self.control_by_id("Combo05")
        selected_text = selected_data.window_text().strip() if selected_data else ""
        if "".join(selected_text.split()).casefold() != "".join(report_type.split()).casefold():
            raise RuntimeError(
                f"ประเภทข้อมูลไม่ตรง: ต้องการ {report_type}, แต่หน้าจอแสดง {selected_text or 'ไม่พบค่า'}"
            )
        self.log("ตั้งค่ารายงาน", "OK", f"ประเภทข้อมูล: {selected_text}")
        self.choose_calendar_date("DateTimeEdit01", self.start_date)
        self.choose_calendar_date("DateTimeEdit02", self.end_date)
        return self.preview_and_save_report(organization, report_type, output, started_at)

    def preview_and_save_report(self, organization: str, report_type: str,
                                output: Path, started_at: datetime) -> Path:
        preview = self.report_preview()
        if not preview:
            raise RuntimeError("ไม่พบหน้ารายงาน")
        preview_button = next((e for e in preview.descendants(control_type="Button")
                               if e.window_text() == "Preview"), None)
        if not preview_button:
            raise RuntimeError("ไม่พบปุ่ม Preview")
        preview_button.click_input()
        deadline = time.monotonic() + self.args.wait_seconds
        while time.monotonic() < deadline:
            preview = self.report_preview()
            page = next((e.window_text() for e in preview.descendants(control_type="Button")
                         if e.window_text().startswith("Page ")), "") if preview else ""
            if page and "Page 0 of 0" not in page:
                break
            time.sleep(.5)
        else:
            raise RuntimeError("Preview รายงานไม่เสร็จภายในเวลาที่กำหนด")
        export = next((e for e in preview.descendants(control_type="Button")
                       if e.window_text() == "Export"), None)
        if not export:
            raise RuntimeError("ไม่พบปุ่ม Export")
        rect = export.rectangle()
        self.click_owned(rect.right - 5, int(rect.mid_point().y))
        time.sleep(.4)
        excel_data = next((e for w in self.windows() for e in w.descendants(control_type="MenuItem")
                           if e.window_text() == "Excel data" and e.is_visible()), None)
        if not excel_data:
            raise RuntimeError("ไม่พบเมนู Export > Excel data")
        excel_data.click_input()
        deadline = time.monotonic() + self.args.wait_seconds
        save_dialog = None
        while time.monotonic() < deadline:
            preview = self.report_preview()
            save_dialog = next((e for e in preview.descendants(control_type="Window")
                                if e.window_text() == "Save As"), None) if preview else None
            if save_dialog:
                break
            time.sleep(.3)
        if not save_dialog:
            raise RuntimeError("ไม่พบหน้าต่าง Save As")
        filename = next((e for e in save_dialog.descendants(control_type="Edit")
                         if e.element_info.automation_id == "1001"), None)
        save = next((e for e in save_dialog.descendants(control_type="Button")
                     if e.element_info.automation_id == "1"), None)
        if not filename or not save:
            raise RuntimeError("ไม่พบช่องชื่อไฟล์หรือปุ่ม Save")
        if output.exists():
            output.unlink()
        filename.set_edit_text(str(output.resolve()))
        save.click_input()
        deadline = time.monotonic() + self.args.wait_seconds
        while time.monotonic() < deadline:
            if output.is_file() and output.stat().st_size > 0:
                alert = next((w for w in self.windows()
                              if w.element_info.automation_id == "frmAlert"), None)
                if alert:
                    alert.close()
                    time.sleep(.4)
                self.record_report_export(
                    organization, report_type, output, "สำเร็จ",
                    started_at=started_at, finished_at=datetime.now(BANGKOK_TZ),
                )
                self.done(f"บันทึก {output}")
                return output
            time.sleep(.4)
        raise RuntimeError(f"Export แล้วแต่ไม่พบไฟล์: {output}")

    def navigate_approve_center(self) -> None:
        self.start("เปิดเมนู Approve Center")
        alert = next((w for w in self.windows()
                      if w.element_info.automation_id == "frmAlert"), None)
        if alert:
            alert.close()
            time.sleep(.4)
        deadline = time.monotonic() + self.args.wait_seconds
        main = self.largest_window()
        while main is None and time.monotonic() < deadline:
            time.sleep(.3)
            main = self.largest_window()
        if main is None:
            raise RuntimeError("ไม่พบหน้าหลัก HRMI หลังเลือกองค์กร")
        try:
            main.maximize()
            time.sleep(.3)
        except Exception:
            pass
        button = next((e for e in main.descendants(control_type="Button")
                       if e.window_text() == "Approve Center" and e.is_visible()), None)
        if not button:
            raise RuntimeError("ไม่พบเมนู Approve Center")
        button.click_input()
        if not self.wait_text("Approve Center Flow", 8):
            raise RuntimeError("เปิด Approve Center แล้วแต่ไม่พบหน้า Approve Center Flow")
        self.done("เปิด Approve Center แล้ว")

    def open_approval_leave_report(self) -> None:
        alert = next((w for w in self.windows()
                      if w.element_info.automation_id == "frmAlert"), None)
        if alert:
            alert.close()
            time.sleep(.4)
        preview = self.report_preview()
        if preview:
            if any(e.element_info.automation_id == "frmReportRangeForm"
                   for e in preview.descendants()):
                return
            range_button = next((e for e in preview.descendants(control_type="Button")
                                 if e.window_text() == "Range"), None)
            if range_button:
                range_button.click_input()
        else:
            main = self.largest_window()
            self.activate(main)
            tree = next((e for e in main.descendants(control_type="Tree") if e.is_visible()), None)
            if not tree:
                raise RuntimeError("ไม่พบรายการเมนู Approve Center")
            point = tree.rectangle().mid_point()
            mouse.scroll(coords=(int(point.x), int(point.y)), wheel_dist=100)
            time.sleep(.5)
            item = None
            for _ in range(40):
                main = self.largest_window()
                item = next((e for e in main.descendants(control_type="TreeItem")
                             if e.window_text() == "Node27" and e.is_visible()), None)
                if item and item.rectangle().bottom <= tree.rectangle().bottom:
                    break
                line_down = next((e for e in tree.descendants(control_type="Button")
                                  if e.window_text() == "Line Down"), None)
                if not line_down:
                    break
                line_down.click_input()
                time.sleep(.08)
            if not item:
                raise RuntimeError("ไม่พบเมนูรายงานข้อมูลการอนุมัติลา (Node27)")
            if item.rectangle().bottom > tree.rectangle().bottom:
                raise RuntimeError("เมนูรายงานข้อมูลการอนุมัติลาอยู่นอกพื้นที่คลิก")
            item.double_click_input()
        deadline = time.monotonic() + max(self.args.wait_seconds, 90)
        while time.monotonic() < deadline:
            preview = self.report_preview()
            if preview and any(e.element_info.automation_id == "frmReportRangeForm"
                               for e in preview.descendants()):
                return
            time.sleep(.3)
        raise RuntimeError("ไม่พบหน้ากำหนดช่วงรายงานข้อมูลการอนุมัติลา")

    def clear_report_date(self, automation_id: str) -> None:
        control = self.control_by_id(automation_id)
        if not control:
            raise RuntimeError(f"ไม่พบช่องวันที่ {automation_id}")
        control.click_input()
        keyboard.send_keys("^a{BACKSPACE}{TAB}")
        time.sleep(.2)
        actual = self.control_by_id(automation_id)
        if not actual or actual.window_text().strip():
            shown = actual.window_text() if actual else "ไม่พบช่อง"
            raise RuntimeError(f"ล้างวันที่ {automation_id} ไม่สำเร็จ: {shown}")

    def export_approval_leave_report(self, organization: str) -> Path:
        report_type = "อนุมัติลา"
        output = self.report_file(organization, report_type)
        output.parent.mkdir(parents=True, exist_ok=True)
        started_at = datetime.now(BANGKOK_TZ)
        self.start("Export รายงาน", f"{organization} {report_type} -> {output}")
        self.open_approval_leave_report()
        preview = self.report_preview()
        if not preview:
            raise RuntimeError("ไม่พบหน้ารายงานข้อมูลการอนุมัติลา")
        controls = {e.element_info.automation_id: e for e in preview.descendants()
                    if e.element_info.automation_id}
        if not controls.get("cmbReportList"):
            raise RuntimeError("ไม่พบช่องประเภทรายงานข้อมูลการอนุมัติลา")
        if not all(controls.get(key) for key in
                   ("DateTimeEdit01", "DateTimeEdit02", "DateTimeEdit03", "DateTimeEdit04")):
            raise RuntimeError("ช่องวันที่ของรายงานข้อมูลการอนุมัติลาไม่ครบ")
        self.choose_combo_text(controls["cmbReportList"], "เรียงตามหน่วยงาน")
        self.clear_report_date("DateTimeEdit01")
        self.clear_report_date("DateTimeEdit02")
        self.choose_calendar_date("DateTimeEdit03", self.start_date)
        self.choose_calendar_date("DateTimeEdit04", self.end_date)
        return self.preview_and_save_report(organization, report_type, output, started_at)

    def close_report_preview(self) -> None:
        preview = self.report_preview()
        if not preview:
            return
        close = next((e for e in preview.descendants(control_type="Button")
                      if e.window_text() == "Close" and e.rectangle().top < 80), None)
        if close:
            close.click_input()
            time.sleep(.5)
        else:
            preview.close()

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
        expected = (f"{self.start_date:%d/%m}/{self.start_date.year + 543}  -  "
                    f"{self.end_date:%d/%m}/{self.end_date.year + 543}")
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

    def ensure_organization_picker(self, user: str, password: str) -> None:
        """Move any supported HRMI state to the Organization picker."""
        if self.window("Organization"):
            return
        if self.window("Login"):
            if not self.args.auto_login:
                raise RuntimeError("HRMI อยู่หน้า Login; ใช้ --auto-login หรือ Login เองก่อนรัน")
            self.login(user, password)
            return
        self.close_report_preview()
        if not any(w.element_info.automation_id == "frmFrame" for w in self.windows()):
            raise RuntimeError("ไม่พบหน้าหลัก HRMI สำหรับกลับไปหน้าเลือกองค์กร")
        self.log("เตรียมหน้าองค์กร", "INFO", "HRMI เปิดค้างที่หน้าหลัก; กลับหน้าเลือกองค์กรอัตโนมัติ")
        self.return_to_picker(user, password)


def main() -> int:
    parser = argparse.ArgumentParser(description="Run HRMI time processing wizard for all configured organizations")
    parser.add_argument("--exe", type=Path, default=DEFAULT_EXE)
    parser.add_argument("--wait-seconds", type=int, default=30)
    parser.add_argument("--organizations", nargs="+", default=list(ORGANIZATIONS), choices=list(ORGANIZATIONS))
    parser.add_argument("--log-path", type=Path, default=ROOT / "logs")
    parser.add_argument("--calibrate-time-attendance", action="store_true")
    parser.add_argument("--auto-login", action="store_true")
    parser.add_argument("--scheduled", action="store_true",
                        help="unattended mode: resume missing organizations and never prompt")
    parser.add_argument("--force", action="store_true",
                        help="process all selected organizations even if already completed")
    parser.add_argument("--start-date", type=parse_cli_date,
                        help="manual range start: YYYY-MM-DD or DD/MM/YYYY")
    parser.add_argument("--end-date", type=parse_cli_date,
                        help="manual range end: YYYY-MM-DD or DD/MM/YYYY")
    parser.add_argument("--stop-before-next", action="store_true",
                        help="open and set dates only; do not advance the wizard")
    parser.add_argument("--export-reports", action="store_true",
                        help="export ขาด and สาย reports to Excel after processing")
    parser.add_argument("--report-only", action="store_true",
                        help="skip time processing and export only missing reports")
    parser.add_argument("--export-path", type=Path, default=ROOT / "exports",
                        help="root folder for exported Excel reports")
    args = parser.parse_args()
    if bool(args.start_date) != bool(args.end_date):
        parser.error("ต้องระบุ --start-date และ --end-date พร้อมกัน")
    if args.start_date and args.end_date:
        if args.scheduled:
            parser.error("--scheduled ใช้วันที่ย้อนหลัง 2 วันและใช้ร่วมกับ custom range ไม่ได้")
        if args.end_date < args.start_date:
            parser.error("--end-date ต้องไม่น้อยกว่า --start-date")
        if (args.start_date.year, args.start_date.month) != (args.end_date.year, args.end_date.month):
            parser.error("ช่วงวันที่ต้องอยู่ภายในเดือนเดียวกัน เพราะ HRMI เลือกเดือนได้ครั้งละหนึ่งเดือน")
    if args.scheduled:
        args.auto_login = True
    if args.report_only:
        args.export_reports = True
    load_env()
    user = os.getenv("HRMI_USER") or os.getenv("user", "")
    password = os.getenv("HRMI_PASSWORD") or os.getenv("password", "")
    if args.auto_login and (not user or not password):
        parser.error("ตั้ง HRMI_USER และ HRMI_PASSWORD ก่อนใช้ --auto-login")
    if args.auto_login and (user.strip().casefold() in {"ชื่อผู้ใช้", "username", "user name", "ใส่ชื่อผู้ใช้"}
                            or password.strip().casefold() in {"รหัสผ่าน", "password", "ใส่รหัสผ่าน"}):
        parser.error("ข้อมูล Login ยังเป็นข้อความตัวอย่าง")
    with SingleInstance() as instance:
        if instance.already_running:
            print("ข้าม: มี HRMI AUTO กำลังทำงานอยู่แล้ว")
            return 0
        hrmi = Hrmi(args)
        try:
            hrmi.log("เริ่มต้น", "INFO", f"Log: {hrmi.log_path}")
            if args.scheduled:
                rule = "วันอังคาร: เสาร์-อาทิตย์ล่าสุด" if hrmi.start_date != hrmi.end_date else "ย้อนหลัง 2 วัน"
                hrmi.log("ช่วงวันที่อัตโนมัติ", "INFO",
                         f"{thai_date_text(hrmi.start_date)} - {thai_date_text(hrmi.end_date)} ({rule})")
            completed = completed_organizations(hrmi.timing_json, hrmi.start_date, hrmi.end_date)
            processing_selected = [] if args.report_only else organizations_to_run(
                args.organizations, completed, scheduled=args.scheduled, force=args.force,
            )
            report_selected = [
                org for org in args.organizations
                if args.export_reports and (args.force or any(
                    not hrmi.report_is_complete(org, report_type) for report_type in REPORT_TYPES
                ))
            ]
            selected = [org for org in args.organizations
                        if org in processing_selected or org in report_selected]
            skipped = [org for org in args.organizations if org not in processing_selected]
            if skipped:
                hrmi.log("ตรวจรอบเดิม", "INFO", f"ข้ามการประมวลผลสาขาที่สำเร็จแล้ว: {', '.join(skipped)}")
            if not selected:
                hrmi.log("จบสคริปต์", "INFO", "ไม่มีงานประมวลผลหรือรายงานที่ค้างอยู่")
                return 0

            hrmi.connect()
            hrmi.ensure_organization_picker(user, password)
            if not hrmi.window("Organization"):
                raise RuntimeError("ไม่พบหน้าเลือกองค์กร; ตรวจยืนยันองค์กรปัจจุบันไม่ได้")
            for index, org in enumerate(selected):
                org_started = datetime.now(BANGKOK_TZ)
                processing_recorded = False
                try:
                    hrmi.select_org(org)
                    if org in processing_selected:
                        hrmi.open_processing()
                        if not args.stop_before_next:
                            hrmi.advance_wizard()
                        message = (COMPLETED_MESSAGE if not args.stop_before_next
                                   else "เปิดหน้าประมวลผลและตั้งวันที่แล้ว")
                        hrmi.log("ประมวลผลองค์กร", "OK", message)
                        hrmi.record_timing(org, org_started, datetime.now(BANGKOK_TZ), "สำเร็จ", message)
                        processing_recorded = True
                    if args.export_reports and not args.stop_before_next:
                        hrmi.navigate()
                        for report_type in ABSENCE_REPORT_TYPES:
                            if args.force or not hrmi.report_is_complete(org, report_type):
                                try:
                                    hrmi.export_absence_report(org, report_type)
                                except Exception as exc:
                                    hrmi.record_report_export(
                                        org, report_type, hrmi.report_file(org, report_type),
                                        "ผิดพลาด", str(exc), finished_at=datetime.now(BANGKOK_TZ),
                                    )
                                    raise
                            else:
                                hrmi.log("Export รายงาน", "INFO", f"ข้ามไฟล์ที่มีแล้ว: {org} {report_type}")
                        hrmi.close_report_preview()
                except Exception as exc:
                    if org in processing_selected and not processing_recorded:
                        hrmi.record_timing(org, org_started, datetime.now(BANGKOK_TZ), "ผิดพลาด", str(exc))
                    raise
                if index < len(selected) - 1:
                    hrmi.return_to_picker(user, password)
            hrmi.log("จบสคริปต์", "INFO", f"เสร็จ {len(selected)} องค์กร")
            return 0
        except Exception as exc:
            hrmi.log(hrmi.step, "FATAL", str(exc))
            print(f"หยุด: {exc}", file=sys.stderr)
            return 1
        finally:
            hrmi.close()


if __name__ == "__main__":
    raise SystemExit(main())
