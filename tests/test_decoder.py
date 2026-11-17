import csv
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import load_workbook

import decoder


DAY1, DAY2, FILM = "2026-10-01", "2026-10-02", "00103542025"
PREFIX = f"{FILM}_{DAY1}_{DAY2}"
FIELDS = ("date", "cinema_code", "cinema_name", "movie_code", "status", "sales_fen")


def item(code, day, status, cents=""):
    return (day, code, f"影院 {code}", FILM, status, cents)


class DecoderTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.file = self.root / f"{PREFIX}_cinema_days.csv"
        self.xlsx = self.root / "逐影院票房.xlsx"

    def write(self, rows):
        with self.file.open("w", encoding="utf-8-sig", newline="") as handle:
            csv.writer(handle).writerows([FIELDS, *rows])

    def read(self):
        book = load_workbook(self.xlsx)
        self.addCleanup(book.close)
        return book

    def convert(self):
        return decoder.convert(decoder.discover(self.file)[0], self.xlsx)

    def test_each_cinema_owns_its_own_amount_and_zero_requires_verification(self):
        self.write([
            item("00112233", DAY1, "with_movie", "12345"), item("00112233", DAY2, "with_movie", "2500"),
            item("00112234", DAY1, "with_movie", "35000"), item("00112234", DAY2, "deferred_204", "1000"),
            item("00112235", DAY1, "without_movie", "0"), item("00112235", DAY2, "without_movie", "0"),
            item("00112236", DAY1, "error"), item("00112236", DAY2, "missing"),
        ])
        self.assertEqual(self.convert(), 4)
        wb = self.read()
        self.assertEqual(wb.sheetnames, ["逐影院票房"])
        sheet = wb.active
        self.assertEqual(sheet.auto_filter.ref, "A6:I10")
        self.assertEqual(sheet.freeze_panes, "C7")
        self.assertEqual(sheet["A7"].data_type, "s")
        rows = list(sheet.iter_rows(min_row=7, values_only=True))
        self.assertEqual([row[0] for row in rows], ["00112234", "00112233", "00112235", "00112236"])
        self.assertEqual(rows[0][2:8], (350, None, "部分日期待核查", 1, 1, 1))
        self.assertEqual(rows[1][2:8], (148.45, 148.45, "已核查", 2, 0, 0))
        self.assertEqual(rows[2][2:5], (0, 0, "已核查，无该片"))
        self.assertEqual(rows[3][2:8], (None, None, "待核查", 0, 2, 0))

    def test_missing_row_and_named_interval_edges_still_count_as_unknown(self):
        self.write([item("00112233", DAY1, "with_movie", "200")])
        self.convert()
        sheet = self.read().active
        self.assertEqual(sheet["C7"].value, 2)
        self.assertIsNone(sheet["D7"].value)
        self.assertEqual(sheet["G7"].value, 1)

    def test_duplicates_do_not_add_and_conflicts_do_not_overwrite_existing_report(self):
        row = item("00112233", DAY1, "with_movie", "123")
        self.write([row, row, item("00112233", DAY2, "without_movie", "0")])
        self.assertEqual(self.convert(), 1)
        wb = load_workbook(self.xlsx)
        self.assertEqual(wb.active["D7"].value, 1.23)
        wb.close()
        saved = self.xlsx.read_bytes()
        self.write([row, (*row[:-1], "900")])
        with self.assertRaisesRegex(ValueError, "冲突"):
            self.convert()
        self.assertEqual(self.xlsx.read_bytes(), saved)

    def test_zip_or_directory_ignores_roster_audit_and_json(self):
        self.write([item("00112233", DAY1, "with_movie", "600")])
        roster = self.root / "cinema_roster.csv"
        roster.write_text("date,cinema_code\n", encoding="utf-8")
        audit = self.root / f"{PREFIX}_daily_audit.csv"
        audit.write_text("date,target_count\n", encoding="utf-8")
        archive = self.root / "results.zip"
        with ZipFile(archive, 'w') as z:
            for path in (self.file, roster, audit):
                z.write(path, "output/national_history/" + path.name)
        self.assertEqual(decoder.discover(roster), [])
        self.assertEqual(decoder.discover(audit), [])
        self.assertEqual(len(decoder.discover(self.root)), 1)
        self.assertEqual(len(decoder.discover(archive)), 1)
        self.assertEqual(decoder.convert(decoder.discover(archive)[0], self.xlsx), 1)

    def test_renamed_source_does_not_claim_a_complete_period(self):
        self.file = self.root / "改过名字.csv"
        self.write([item("00112233", DAY1, "with_movie", "500")])
        self.convert()
        sheet = self.read().active
        self.assertEqual(sheet["C7"].value, 5)
        self.assertIsNone(sheet["D7"].value)
        self.assertEqual(sheet["E7"].value, "导出日期范围待核实")

    def test_desktop_exports_the_per_cinema_workbook(self):
        from app import App
        self.write([item("00112233", DAY1, "with_movie", "123")])
        app = App()
        try:
            app.withdraw()
            app.load(self.file)
            self.assertIn("已确认票房", app.note.get())
            with patch("app.filedialog.asksaveasfilename", return_value=str(self.xlsx)):
                app.generate.invoke()
            deadline = time.monotonic() + 8
            while app.busy and time.monotonic() < deadline:
                app.update()
                time.sleep(0.02)
            self.assertFalse(app.busy, app.status.get())
            self.assertIn("1 家影院", app.status.get())
            self.assertEqual(self.read().sheetnames, ["逐影院票房"])
        finally:
            app.close()
