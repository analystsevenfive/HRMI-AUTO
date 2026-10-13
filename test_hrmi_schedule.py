import json
import tempfile
import unittest
import uuid
import zipfile
from datetime import date
from pathlib import Path

from hrmi import (Hrmi, SingleInstance, completed_organizations, organizations_to_run,
                  parse_cli_date, scheduled_processing_range)
from rocket75_import import date_folder, report_file, xlsx_has_data


ORGS = ["[BB]", "[BK]", "[NTS]", "[PK]", "[SM]"]
TARGET = date(2026, 10, 10)


class ScheduleTests(unittest.TestCase):
    def timing_file(self, rows):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        path = Path(directory.name) / "timing.json"
        path.write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
        return path

    def successful_row(self, organization, message="ประมวลผลและกด Finish แล้ว"):
        return {
            "target_date": TARGET.isoformat(),
            "organization": organization,
            "status": "สำเร็จ",
            "message": message,
        }

    def test_scheduled_new_day_runs_every_organization(self):
        self.assertEqual(organizations_to_run(ORGS, set(), scheduled=True, force=False), ORGS)

    def test_scheduled_partial_day_resumes_missing_organizations(self):
        completed = {"[BB]", "[BK]"}
        self.assertEqual(
            organizations_to_run(ORGS, completed, scheduled=True, force=False),
            ["[NTS]", "[PK]", "[SM]"],
        )

    def test_scheduled_complete_day_skips_run(self):
        self.assertEqual(organizations_to_run(ORGS, set(ORGS), scheduled=True, force=False), [])

    def test_manual_complete_day_asks_before_rerun(self):
        denied = organizations_to_run(ORGS, set(ORGS), scheduled=False, force=False, input_fn=lambda _: "n")
        accepted = organizations_to_run(ORGS, set(ORGS), scheduled=False, force=False, input_fn=lambda _: "y")
        self.assertEqual(denied, [])
        self.assertEqual(accepted, ORGS)

    def test_force_runs_every_organization(self):
        self.assertEqual(organizations_to_run(ORGS, set(ORGS), scheduled=True, force=True), ORGS)

    def test_only_finished_processing_counts_as_complete(self):
        path = self.timing_file([
            self.successful_row("[BB]"),
            self.successful_row("[BK]", "เปิดหน้าประมวลผลและตั้งวันที่แล้ว"),
            {**self.successful_row("[NTS]"), "status": "ผิดพลาด"},
        ])
        self.assertEqual(completed_organizations(path, TARGET), {"[BB]"})

    def test_custom_range_is_matched_separately(self):
        path = self.timing_file([{
            **self.successful_row("[BB]"),
            "target_start_date": "2026-10-01",
            "target_end_date": "2026-10-05",
        }])
        self.assertEqual(
            completed_organizations(path, date(2026, 10, 1), date(2026, 10, 5)),
            {"[BB]"},
        )
        self.assertEqual(completed_organizations(path, TARGET), set())

    def test_cli_date_accepts_iso_gregorian_and_buddhist_year(self):
        self.assertEqual(parse_cli_date("2026-10-01"), date(2026, 10, 1))
        self.assertEqual(parse_cli_date("01/10/2026"), date(2026, 10, 1))
        self.assertEqual(parse_cli_date("01/10/2569"), date(2026, 10, 1))

    def test_tuesday_processes_latest_saturday_and_sunday(self):
        self.assertEqual(
            scheduled_processing_range(date(2026, 10, 13)),
            (date(2026, 10, 10), date(2026, 10, 11)),
        )

    def test_non_tuesday_processes_one_day_two_days_back(self):
        self.assertEqual(
            scheduled_processing_range(date(2026, 10, 14)),
            (date(2026, 10, 12), date(2026, 10, 12)),
        )

    def test_tuesday_weekend_can_cross_month_boundary(self):
        self.assertEqual(
            scheduled_processing_range(date(2026, 11, 3)),
            (date(2026, 10, 31), date(2026, 11, 1)),
        )

    def test_single_instance_detects_second_process_guard(self):
        name = rf"Local\HRMI_AUTO_TEST_{uuid.uuid4().hex}"
        with SingleInstance(name) as first:
            with SingleInstance(name) as second:
                self.assertFalse(first.already_running)
                self.assertTrue(second.already_running)

    def test_report_file_uses_organization_type_and_date_range(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        hrmi = Hrmi.__new__(Hrmi)
        hrmi.export_path = Path(directory.name)
        hrmi.start_date = date(2026, 10, 1)
        hrmi.end_date = date(2026, 10, 10)
        expected = Path(directory.name) / "2026-10-01_2026-10-10" / "HRMI_BB_ขาด_2026-10-01_2026-10-10.xlsx"
        self.assertEqual(hrmi.report_file("[BB]", "ขาด"), expected)
        self.assertFalse(hrmi.report_is_complete("[BB]", "ขาด"))
        expected.parent.mkdir(parents=True)
        expected.write_bytes(b"xlsx")
        self.assertTrue(hrmi.report_is_complete("[BB]", "ขาด"))

    def test_rocket75_files_are_separated_by_type(self):
        root = Path("exports")
        start = date(2026, 10, 10)
        end = date(2026, 10, 11)
        self.assertEqual(date_folder(start, end), "2026-10-10_2026-10-11")
        self.assertEqual(
            report_file(root, "BB", "ขาด", start, end),
            root / "2026-10-10_2026-10-11" / "HRMI_BB_ขาด_2026-10-10_2026-10-11.xlsx",
        )
        self.assertEqual(
            report_file(root, "BB", "สาย", start, end),
            root / "2026-10-10_2026-10-11" / "HRMI_BB_สาย_2026-10-10_2026-10-11.xlsx",
        )

    def test_rocket75_skips_header_only_workbook(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        path = Path(directory.name) / "report.xlsx"
        header = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                  '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>EmpCode</t></is></c>'
                  '</row></sheetData></worksheet>')
        with zipfile.ZipFile(path, "w") as workbook:
            workbook.writestr("xl/worksheets/sheet1.xml", header)
        self.assertFalse(xlsx_has_data(path))

        data = header.replace('</row></sheetData>',
                              '</row><row r="2"><c r="A2" t="inlineStr"><is><t>0125001</t></is></c>'
                              '</row></sheetData>')
        with zipfile.ZipFile(path, "w") as workbook:
            workbook.writestr("xl/worksheets/sheet1.xml", data)
        self.assertTrue(xlsx_has_data(path))


if __name__ == "__main__":
    unittest.main()
