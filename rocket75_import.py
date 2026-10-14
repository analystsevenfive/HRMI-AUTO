import argparse
import json
import os
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path
from xml.etree import ElementTree

import requests

from hrmi import BANGKOK_TZ, ROOT, load_env, parse_cli_date, scheduled_processing_range


BASE_URL = "https://rocket75.com/hr"
LOGIN_URL = "https://rocket75.com/auth.php"
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")
ENDPOINTS = {
    "ขาด": f"{BASE_URL}/ajax/import_absenteeism/Import.php",
    "สาย": f"{BASE_URL}/ajax/import_checkin/Import.php",
    "ลา": f"{BASE_URL}/ajax/import_leave_work/Import.php",
}
TYPE_ALIASES = {"absence": "ขาด", "late": "สาย", "leave": "ลา",
                "ขาด": "ขาด", "สาย": "สาย", "ลา": "ลา"}
ORGANIZATIONS = ("BB", "BK", "NTS", "PK", "SM")
SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def date_folder(start: date, end: date) -> str:
    return start.isoformat() if start == end else f"{start.isoformat()}_{end.isoformat()}"


def report_file(export_root: Path, organization: str, report_type: str,
                start: date, end: date) -> Path:
    period = date_folder(start, end)
    return export_root / period / f"HRMI_{organization}_{report_type}_{period}.xlsx"


def created_date(path: Path) -> date:
    """Read Windows Created date of a folder in Bangkok time."""
    metadata = path.stat()
    timestamp = getattr(metadata, "st_birthtime", None)
    if timestamp is None:
        if os.name != "nt":
            raise RuntimeError("ระบบนี้ไม่รองรับวันที่ Created ของไฟล์")
        timestamp = metadata.st_ctime
    return datetime.fromtimestamp(timestamp, BANGKOK_TZ).date()


def resolve_today_leave_folder(root: Path, today: date | None = None) -> Path:
    """Select the only direct child folder whose Windows creation date is today."""
    if not root.is_dir():
        raise RuntimeError(f"ไม่พบ Share Center root: {root}")
    target_date = today or datetime.now(BANGKOK_TZ).date()
    matches = sorted(
        (path for path in root.iterdir()
         if path.is_dir() and created_date(path) == target_date),
        key=lambda path: path.name,
    )
    if not matches:
        raise RuntimeError(f"ไม่พบโฟลเดอร์ที่สร้างวันที่ {target_date:%d/%m/%Y} ใน {root}")
    if len(matches) > 1:
        names = ", ".join(path.name for path in matches)
        raise RuntimeError(
            f"พบโฟลเดอร์ที่สร้างวันที่ {target_date:%d/%m/%Y} มากกว่า 1 โฟลเดอร์ใน {root}: {names}"
        )
    return matches[0]


def xlsx_has_data(path: Path) -> bool:
    """Return True when any worksheet contains a non-empty value below row 1."""
    try:
        with zipfile.ZipFile(path) as workbook:
            sheets = sorted(name for name in workbook.namelist()
                            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"))
            for sheet in sheets:
                root = ElementTree.fromstring(workbook.read(sheet))
                for row in root.iter(f"{{{SPREADSHEET_NS}}}row"):
                    row_number = int(row.attrib.get("r", "0") or 0)
                    if row_number <= 1:
                        continue
                    for cell in row.findall(f"{{{SPREADSHEET_NS}}}c"):
                        values = [node.text or "" for node in cell.iter()
                                  if node.tag in {f"{{{SPREADSHEET_NS}}}v",
                                                  f"{{{SPREADSHEET_NS}}}t"}]
                        if any(value.strip() for value in values):
                            return True
    except (OSError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise RuntimeError(f"อ่านไฟล์ Excel ไม่สำเร็จ: {path}") from exc
    return False


def upload_file(session: requests.Session, path: Path, report_type: str,
                key: str, token: str) -> dict:
    with path.open("rb") as stream:
        response = session.post(
            ENDPOINTS[report_type],
            data={"key": key, "token": token},
            files={"uploadfile": (path.name, stream,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            timeout=120,
        )
    response.raise_for_status()
    try:
        result = response.json()
    except requests.JSONDecodeError as exc:
        raise RuntimeError(f"เว็บตอบกลับไม่ใช่ JSON: HTTP {response.status_code}") from exc
    if result.get("result") == 1:
        return result
    if result.get("result") == 0:
        raise RuntimeError(f"เว็บแจ้งว่ามีข้อมูลซ้ำ: {path.name}")
    raise RuntimeError(f"เว็บนำเข้าไม่สำเร็จ: {path.name}; {json.dumps(result, ensure_ascii=False)}")


def login(session: requests.Session, username: str, password: str) -> tuple[str, str]:
    response = session.post(
        LOGIN_URL,
        data={"username": username, "password": password},
        timeout=30,
    )
    response.raise_for_status()
    try:
        result = response.json()
    except requests.JSONDecodeError as exc:
        raise RuntimeError(f"Login ตอบกลับไม่ใช่ JSON: HTTP {response.status_code}") from exc
    if result.get("sing") not in (1, 2):
        raise RuntimeError("Login Rocket75 ไม่สำเร็จ: ตรวจ ROCKET_USER และ ROCKET_PASSWORD")
    key = str(result.get("key") or "").strip()
    token = str(result.get("token") or "").strip()
    if not key or not token:
        raise RuntimeError("Login สำเร็จแต่ Rocket75 ไม่ส่ง key/token กลับมา")
    return key, token


def main() -> int:
    parser = argparse.ArgumentParser(description="Import HRMI absence and late reports into Rocket75")
    parser.add_argument("--start-date", type=parse_cli_date)
    parser.add_argument("--end-date", type=parse_cli_date)
    parser.add_argument("--organizations", nargs="+", default=list(ORGANIZATIONS), choices=ORGANIZATIONS)
    parser.add_argument("--types", nargs="+", default=["absence", "late"], choices=list(TYPE_ALIASES))
    parser.add_argument("--export-path", type=Path, default=ROOT / "exports")
    parser.add_argument("--leave-folder", type=Path,
                        help="Share Center folder containing approved-leave .xlsx files")
    parser.add_argument("--leave-root", type=Path,
                        help="Share Center root; select its child folder created today")
    args = parser.parse_args()
    if bool(args.start_date) != bool(args.end_date):
        parser.error("ต้องระบุ --start-date และ --end-date พร้อมกัน")
    start, end = ((args.start_date, args.end_date) if args.start_date
                  else scheduled_processing_range())

    load_env()
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept": "application/json, text/javascript, */*; q=0.01"})
    username = os.getenv("ROCKET_USER", "").strip()
    password = os.getenv("ROCKET_PASSWORD", "").strip()
    if username and password:
        print("Login Rocket75 และสร้าง token ใหม่", flush=True)
        key, token = login(session, username, password)
        print("Login Rocket75 สำเร็จ", flush=True)
    else:
        key = os.getenv("ROCKET75_KEY", "").strip()
        token = os.getenv("ROCKET75_TOKEN", "").strip()
        if not key or not token:
            print("หยุด: เพิ่ม ROCKET_USER/ROCKET_PASSWORD หรือ ROCKET75_KEY/ROCKET75_TOKEN ใน .env",
                  file=sys.stderr)
            return 2

    selected_types = [TYPE_ALIASES[value] for value in args.types]
    jobs = [(org, kind, report_file(args.export_path, org, kind, start, end))
            for org in args.organizations for kind in selected_types if kind != "ลา"]
    if "ลา" in selected_types:
        if args.leave_folder and args.leave_root:
            parser.error("ระบุได้อย่างใดอย่างหนึ่ง: --leave-folder หรือ --leave-root")
        try:
            leave_folder = (resolve_today_leave_folder(args.leave_root)
                            if args.leave_root else args.leave_folder)
        except RuntimeError as exc:
            parser.error(str(exc))
        if not leave_folder or not leave_folder.is_dir():
            parser.error("--types leave ต้องระบุ --leave-folder หรือ --leave-root ที่มีอยู่จริง")
        if args.leave_root:
            print(f"เลือกโฟลเดอร์ลาของวันนี้: {leave_folder}", flush=True)
        leave_files = sorted(path for path in leave_folder.glob("*.xlsx")
                             if not path.name.startswith("~$"))
        if not leave_files:
            parser.error(f"ไม่พบไฟล์ .xlsx ใน {leave_folder}")
        jobs.extend((path.stem, "ลา", path) for path in leave_files)
    missing = [str(path) for _, _, path in jobs if not path.is_file() or path.stat().st_size == 0]
    if missing:
        print("หยุด: ไม่พบไฟล์:\n" + "\n".join(missing), file=sys.stderr)
        return 2

    ready_jobs = []
    import_date = datetime.now(BANGKOK_TZ).date()
    skipped_created = 0
    skipped_empty = 0
    print(f"นำเข้าเฉพาะโฟลเดอร์ Created วันที่ {import_date:%d/%m/%Y} (เวลาไทย)", flush=True)
    folder_dates = {path.parent: created_date(path.parent) for _, _, path in jobs}
    for org, kind, path in jobs:
        folder_date = folder_dates[path.parent]
        if folder_date != import_date:
            skipped_created += 1
            print(f"ข้าม [{org}] {kind}: โฟลเดอร์ {path.parent} Created {folder_date:%d/%m/%Y}", flush=True)
            continue
        if xlsx_has_data(path):
            ready_jobs.append((org, kind, path))
        else:
            skipped_empty += 1
            print(f"ข้ามไฟล์เปล่า [{org}] {kind}: {path.name}", flush=True)

    for org, kind, path in ready_jobs:
        print(f"Import [{org}] {kind}: {path.name}", flush=True)
        upload_file(session, path, kind, key, token)
        print(f"สำเร็จ [{org}] {kind}", flush=True)
    print(f"เสร็จ: Import {len(ready_jobs)} ไฟล์; ข้ามเพราะวันที่ Created ของโฟลเดอร์ {skipped_created} ไฟล์; ข้ามไฟล์เปล่า {skipped_empty} ไฟล์")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
