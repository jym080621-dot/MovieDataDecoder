import csv
import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import load_workbook

import decoder


DAY = "2026-11-17"
NEXT = "2026-11-18"
PREFIX = f"00102680026_{DAY}_{NEXT}"


class DecoderTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.detail = self.root / f"{PREFIX}_cinema_days.csv"
        self.audit = self.root / f"{PREFIX}_daily_audit.csv"
        self.xlsx = self.root / "阅读版.xlsx"

    @staticmethod
    def write_csv(path, fields, rows):
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(fields)
            writer.writerows(rows)

    def make_details(self, rows=None):
        self.write_csv(self.detail, ("date", "cinema_code", "cinema_name", "movie_code", "status", "sales_fen"), rows or [
            (DAY, "00112233", "影院甲", "00102680026", "with_movie", "12345"),
            (DAY, "00112234", "影院乙", "00102680026", "without_movie", "0"),
            (DAY, "00112235", "影院丙", "00102680026", "deferred_204", "9000"),
            (NEXT, "00112233", "影院甲", "00102680026", "missing", ""),
            (NEXT, "00112234", "影院乙", "00102680026", "error", ""),
            (NEXT, "00112235", "影院丙", "00102680026", "inconsistent", ""),
        ])

    def make_audit(self):
        rows = (dict(date=DAY, target_count=3, observed=3, **{s: int(s == "with_movie") * 3 for s in decoder.STATES}),)
        with self.audit.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
            writer.writeheader()
            writer.writerows(rows)
        return rows

    def workbook(self):
        workbook = load_workbook(self.xlsx, data_only=False)
        self.addCleanup(workbook.close)
        return workbook

    def test_verified_money_and_distinct_unknown_states(self):
        self.make_details()
        self.assertEqual(decoder.convert(decoder.discover(self.detail)[0], self.xlsx), 6)
        wb = self.workbook()
        self.assertEqual(wb.sheetnames, ["结果概览", "每日汇总", "影院明细"])
        detail = wb["影院明细"]
        self.assertEqual(detail["B5"].value, "00112233")
        self.assertEqual(detail["B5"].data_type, "s")
        self.assertEqual([detail[f"D{r}"].value for r in range(5, 11)],
                         ["有目标电影", "无目标电影", "204 待复核", "未采集", "请求异常", "数据不一致"])
        self.assertEqual([detail[f"E{r}"].value for r in range(5, 11)], [123.45, 0, None, None, None, None])
        self.assertEqual(detail.freeze_panes, "C5")
        self.assertEqual(detail.auto_filter.ref, "A4:F10")
        daily = wb["每日汇总"]
        self.assertEqual(daily["C5"].value, 2)
        self.assertAlmostEqual(daily["J5"].value, 2 / 3)
        self.assertEqual(daily["K5"].value, 123.45)
        self.assertIsNone(daily["K6"].value)

    def test_zip_companion_audit_keeps_stale_counts_separate(self):
        self.make_details()
        self.make_audit()
        archive = self.root / "national_history.zip"
        with ZipFile(archive, "w") as zip_file:
            for rowfile in (self.detail, self.audit):
                zip_file.write(rowfile, f"output/national_history/{rowfile.name}")
        dataset, audit = decoder.discover(archive)
        self.assertEqual(audit.kind, "audit")
        self.assertIn("output/national_history/", dataset.audit.member)
        decoder.convert(dataset, self.xlsx, only_movie=True)
        wb = self.workbook()
        self.assertEqual(wb["每日汇总"]["C5"].value, 2)
        self.assertEqual(wb["审计对照"]["E5"].value, 3)
        self.assertEqual(wb["审计对照"]["F5"].value, "状态计数不一致")
        self.assertEqual(wb["审计对照"]["F6"].value, "缺少对应日期")
        self.assertEqual(wb["影院明细"].max_row, 5)

    def test_audit_only_from_csv_or_json_leaves_money_empty(self):
        rows = self.make_audit()
        summary = self.root / f"{PREFIX}_summary.json"
        summary.write_text(json.dumps({"days": rows}), encoding="utf-8")
        for source in (self.audit, summary):
            with self.subTest(source=source):
                workbook_path = self.root / (source.stem + ".xlsx")
                decoder.convert(decoder.discover(source)[0], workbook_path)
                wb = load_workbook(workbook_path)
                try:
                    self.assertEqual(wb.sheetnames, ["结果概览", "每日汇总"])
                    self.assertIsNone(wb["每日汇总"]["K5"].value)
                    self.assertEqual(wb["每日汇总"]["C5"].value, 3)
                finally:
                    wb.close()

    def test_duplicates_merge_but_conflicting_amount_preserves_existing_output(self):
        entry = (DAY, "00112233", "影院甲", "00102680026", "with_movie", "12345")
        self.make_details([entry, entry])
        self.assertEqual(decoder.convert(decoder.discover(self.detail)[0], self.xlsx), 1)
        previous = self.xlsx.read_bytes()
        self.make_details([entry, (*entry[:-1], "888")])
        with self.assertRaisesRegex(ValueError, "冲突记录"):
            decoder.convert(decoder.discover(self.detail)[0], self.xlsx)
        self.assertEqual(self.xlsx.read_bytes(), previous)

    def test_details_split_at_excel_row_limit(self):
        self.make_details()
        with patch.object(decoder, "MAX_ROWS", 7):
            decoder.convert(decoder.discover(self.detail)[0], self.xlsx)
        wb = self.workbook()
        self.assertEqual(wb["影院明细"].max_row, 7)
        self.assertEqual(wb["影院明细 2"].max_row, 7)
        self.assertEqual(wb["影院明细 2"]["D5"].value, "未采集")

    def test_generic_csv_preserves_leading_zeros_and_literal_text(self):
        source = self.root / "影院名录.csv"
        self.write_csv(source, ("date", "cinema_code", "cinema_name", "national_roster_coverage"),
                       [(DAY, "00112233", "=1+1", "unverified")])
        decoder.convert(decoder.discover(source)[0], self.xlsx)
        sheet = self.workbook()["数据表"]
        self.assertEqual(sheet["B4"].value, "影院编码")
        self.assertEqual(sheet["B5"].value, "00112233")
        self.assertEqual(sheet["C5"].value, "=1+1")
        self.assertEqual(sheet["C5"].data_type, "s")
        self.assertEqual(sheet["D5"].value, "未核实")

    def test_desktop_flow_can_pick_and_convert(self):
        from app import App
        self.make_details()
        app = App()
        try:
            app.withdraw()
            app.load(self.detail)
            self.assertIn("影院明细", app.note.get())
            with patch("app.filedialog.asksaveasfilename", return_value=str(self.xlsx)):
                app.generate.invoke()
            deadline = time.monotonic() + 8
            while app.busy and time.monotonic() < deadline:
                app.update()
                time.sleep(0.02)
            self.assertFalse(app.busy, app.status.get())
            self.assertIn("处理完成", app.status.get())
            self.assertTrue(self.xlsx.is_file())
        finally:
            app.close()


if __name__ == "__main__":
    unittest.main()
