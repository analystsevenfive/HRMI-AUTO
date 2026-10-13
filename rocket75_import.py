import argparse
import json
import os
import sys
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree

import requests

from hrmi import ROOT, load_env, parse_cli_date, scheduled_processing_range


BASE_URL = "https://rocket75.com/hr"
LOGIN_URL = "https://rocket75.com/auth.php"
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")
ENDPOINTS = {
    "ขาด": f"{BASE_URL}/ajax/import_absenteeism/Import.php",
    "สาย": f"{BASE_URL}/ajax/import_checkin/Import.php",
}
TYPE_ALIASES = {"absence": "ขาด", "late": "สาย", "ขาด": "ขาด", "สาย": "สาย"}
ORGANIZATIONS = ("BB", "BK", "NTS", "PK", "SM")
SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def date_folder(start: date, end: date) -> str:
    return start.isoformat() if start == end else f"{start.isoformat()}_{end.isoformat()}"


def report_file(export_root: Path, organization: str, report_type: str,
                start: date, end: date) -> Path:
    period = date_folder(start, end)
    return export_root / period / f"HRMI_{organization}_{report_type}_{period}.xlsx"


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
            for org in args.organizations for kind in selected_types]
    missing = [str(path) for _, _, path in jobs if not path.is_file() or path.stat().st_size == 0]
    if missing:
        print("หยุด: ไม่พบไฟล์:\n" + "\n".join(missing), file=sys.stderr)
        return 2

    ready_jobs = []
    for org, kind, path in jobs:
        if xlsx_has_data(path):
            ready_jobs.append((org, kind, path))
        else:
            print(f"ข้ามไฟล์เปล่า [{org}] {kind}: {path.name}", flush=True)

    for org, kind, path in ready_jobs:
        print(f"Import [{org}] {kind}: {path.name}", flush=True)
        upload_file(session, path, kind, key, token)
        print(f"สำเร็จ [{org}] {kind}", flush=True)
    print(f"เสร็จ: Import {len(ready_jobs)} ไฟล์; ข้ามไฟล์เปล่า {len(jobs) - len(ready_jobs)} ไฟล์")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
