"""Backfill timing workbook for an HRMI run started before timing support loaded."""

from __future__ import annotations

import csv
import json
import subprocess
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import psutil


ROOT = Path(__file__).resolve().parent
BANGKOK_TZ = timezone(timedelta(hours=7))


def read_entries(log_path: Path) -> list[dict[str, str]]:
    try:
        with log_path.open("r", encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))
    except (OSError, csv.Error):
        return []


def sync(log_path: Path) -> bool:
    entries = read_entries(log_path)
    starts: dict[str, datetime] = {}
    completed: list[tuple[str, datetime, datetime, str]] = []
    finished_run = False
    for entry in entries:
        timestamp = datetime.strptime(entry["Timestamp"], "%Y-%m-%d %H:%M:%S.%f").replace(tzinfo=BANGKOK_TZ)
        org = entry.get("Organization", "")
        if entry["Step"] == "เลือกองค์กร" and entry["Status"] == "START" and org:
            starts[org] = timestamp
        elif entry["Step"] == "ประมวลผลองค์กร" and entry["Status"] == "OK" and org in starts:
            completed.append((org, starts[org], timestamp, entry.get("Message", "")))
        elif entry["Step"] == "จบสคริปต์":
            finished_run = True

    timing_path = ROOT / "logs" / "HRMI-processing-times.json"
    rows = []
    if timing_path.exists():
        try:
            rows = json.loads(timing_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            rows = []
    run_id = log_path.stem.removeprefix("HRMI-")
    existing = {(row["run_id"], row["organization"]) for row in rows}
    changed = False
    target_date = (datetime.now(BANGKOK_TZ).date() - timedelta(days=2)).isoformat()
    for org, started, finished, message in completed:
        if (run_id, org) in existing:
            continue
        rows.append({
            "run_id": run_id,
            "organization": org,
            "target_date": target_date,
            "started_at": started.isoformat(),
            "finished_at": finished.isoformat(),
            "duration_seconds": round((finished - started).total_seconds(), 3),
            "status": "สำเร็จ",
            "message": message,
        })
        changed = True
    if changed:
        timing_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        subprocess.run(
            ["node", str(ROOT / "build_processing_times.mjs"), str(timing_path),
             str(ROOT / "HRMI-processing-times.xlsx")],
            cwd=ROOT,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    return finished_run


def main() -> int:
    if len(sys.argv) != 3:
        raise SystemExit("usage: monitor_current_timing.py <log.csv> <hrmi-python-pid>")
    log_path = Path(sys.argv[1]).resolve()
    process_id = int(sys.argv[2])
    while True:
        if sync(log_path):
            return 0
        if not psutil.pid_exists(process_id):
            sync(log_path)
            return 0
        time.sleep(5)


if __name__ == "__main__":
    raise SystemExit(main())
