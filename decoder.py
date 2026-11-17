"""Export per-cinema movie box office totals and daily detail from MovieData."""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import math
import os
import re
import sqlite3
import tempfile
import unicodedata
from contextlib import closing, contextmanager
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

import xlsxwriter


VERSION = "1.2.1"
DETAIL_SUFFIX = "_cinema_days.csv"
DETAIL_FIELDS = {"date", "cinema_code", "movie_code", "status", "sales_fen"}
FILENAME = re.compile(r"([0-9]{1,30})_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})_cinema_days\.csv", re.I)
VERIFIED = {"with_movie", "without_movie"}
EMPTY = {"empty_204", "deferred_204"}
STATUS_LABELS = {
    "with_movie": "已核查，有该片",
    "without_movie": "已核查，无该片",
    "empty_204": "HTTP 204 待复核",
    "deferred_204": "HTTP 204 待复核",
    "missing": "未采集",
    "inconsistent": "数据不一致，待复核",
    "error": "采集异常",
}
MAX_CENTS = 999_999_999_999_999
MAX_ROWS = 1_048_576
HEADER_ROW = 5
NO_DETAILS = "请选择 MovieData 的影院明细（*_cinema_days.csv）。在全国历史页面点「导出当前 CSV / 审计」可生成。"


@dataclass(frozen=True)
class Source:
    path: Path
    member: str = ""

    @property
    def name(self):
        return PurePosixPath(self.member).name if self.member else self.path.name

    @property
    def location(self):
        return f"{self.path.name} / {self.member}" if self.member else self.path.name

    @property
    def filename(self):
        stem = Path(self.name).stem.removesuffix("_cinema_days")
        stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem).strip(" .")[:100] or "MovieData"
        return f"{stem}_逐影院票房.xlsx"

    @contextmanager
    def open(self):
        if self.member:
            with ZipFile(self.path) as archive, archive.open(self.member) as stream:
                with io.TextIOWrapper(stream, encoding="utf-8-sig", newline="") as handle:
                    yield handle
        else:
            with self.path.open(encoding="utf-8-sig", newline="") as handle:
                yield handle


@dataclass
class Cinema:
    verified: int = 0
    with_movie: int = 0
    deferred: int = 0
    cents: int = 0
    name: str = ""
    name_day: dt.date = dt.date.min


@dataclass
class Report:
    movie: str
    start: dt.date
    end: dt.date
    cinemas: dict[str, Cinema]
    inferred_period: bool

    @property
    def days(self):
        return (self.end - self.start).days + 1


def discover(path):
    path = Path(path).resolve()
    if path.is_dir():
        folders = (path, path / "national_history", path / "output" / "national_history")
        sources = {Source(p) for folder in folders if folder.is_dir() for p in folder.iterdir()
                   if p.is_file() and p.name.lower().endswith(DETAIL_SUFFIX)}
    elif path.is_file() and path.suffix.lower() == ".zip":
        with ZipFile(path) as archive:
            names = [info.filename for info in archive.infolist()
                     if not info.is_dir() and info.filename.lower().endswith(DETAIL_SUFFIX)]
        if len(names) != len(set(names)):
            raise ValueError("压缩包里有同名影院明细，请解压后单独选择。")
        sources = {Source(path, name) for name in names}
    elif path.is_file() and path.suffix.lower() == ".csv":
        source = Source(path)
        with source.open() as handle:
            fields = [name.strip() for name in next(csv.reader(handle, strict=True), [])]
        sources = {source} if DETAIL_FIELDS <= set(fields) else set()
    else:
        sources = set()
    return sorted(sources, key=lambda source: (source.location, str(source.path)))


def _day(value):
    day = dt.date.fromisoformat(value)
    if day.isoformat() != value:
        raise ValueError("日期须为 YYYY-MM-DD。")
    return day


def _known_cents(state, value):
    if state not in VERIFIED:
        return None
    if state == "without_movie" and not value:
        return 0
    if not re.fullmatch(r"[0-9]{1,15}", value):
        raise ValueError("已核查片目的票房应为非负整数分。")
    amount = int(value)
    if state == "without_movie" and amount:
        raise ValueError("无目标电影的行包含非零票房。")
    return amount


def _read_report(conn, source, emit):
    name = FILENAME.fullmatch(source.name)
    movie, first, last = "", None, None
    if name:
        movie, start, end = name.groups()
        first, last = _day(start), _day(end)
        if last < first:
            raise ValueError("导出文件名中的结束日期早于开始日期。")

    # A duplicate key can never increase the box office or the verified day count.
    conn.execute("CREATE TABLE seen(day TEXT, cinema TEXT, status TEXT, cents INTEGER, name TEXT, PRIMARY KEY(day,cinema)) WITHOUT ROWID")
    cinemas = {}
    earliest = latest = None
    with source.open() as handle:
        reader = csv.reader(handle, strict=True)
        fields = [field.strip() for field in next(reader, [])]
        if not DETAIL_FIELDS <= set(fields) or len(set(fields)) != len(fields):
            raise ValueError(NO_DETAILS)
        for values in reader:
            if not values:
                continue
            try:
                if len(values) != len(fields):
                    raise ValueError("数据列数与表头不同。")
                row = dict(zip(fields, values))
                day = _day(row["date"].strip())
                code, film, status, raw_fen = (row[key].strip() for key in ("cinema_code", "movie_code", "status", "sales_fen"))
                if not re.fullmatch(r"[0-9]{6,10}", code) or not re.fullmatch(r"[0-9]{1,30}", film) or not status:
                    raise ValueError("影院编码、电影编码或状态无效。")
                if movie and film != movie:
                    raise ValueError("电影编码不一致。每次请导入一部电影的明细。")
                movie = film
                if name and not first <= day <= last:
                    raise ValueError("日期超出了文件名标明的导出区间。")
                cents = _known_cents(status, raw_fen)
                key = day.isoformat(), code
                label = row.get("cinema_name", "").strip()
                inserted = conn.execute("INSERT OR IGNORE INTO seen VALUES (?,?,?,?,?)", (*key, status, cents, label))
                if not inserted.rowcount:
                    old = conn.execute("SELECT status,cents FROM seen WHERE day=? AND cinema=?", key).fetchone()
                    if old != (status, cents):
                        raise ValueError(f"影院 {code} 当日的票房或状态冲突。请重新从 MovieData 导出。")
                    continue
                if earliest is None or day < earliest:
                    earliest = day
                if latest is None or day > latest:
                    latest = day
                if code not in cinemas:
                    cinemas[code] = Cinema()
                cinema = cinemas[code]
                if label and day > cinema.name_day:
                    cinema.name, cinema.name_day = label, day
                if cents is not None:
                    cinema.cents += cents
                    cinema.verified += 1
                    cinema.with_movie += int(status == "with_movie")
                elif status in EMPTY:
                    cinema.deferred += 1
                if reader.line_num % 25_000 == 0:
                    emit(f"已读取 {reader.line_num:,} 行影院明细…")
            except ValueError as exc:
                raise ValueError(f"{source.name} 第 {reader.line_num:,} 行：{exc}") from exc
    if not cinemas:
        raise ValueError("这份明细没有影院记录。请先在 MovieData 采集并导出影院明细。")
    return Report(movie, first or earliest, last or latest, cinemas, not bool(name))


def _report_styles(book):
    base = {"font_name": "Microsoft YaHei", "font_size": 10, "font_color": "#27394B", "valign": "vcenter"}
    styles = {
        "title": book.add_format({**base, "font_size": 16, "bold": True, "font_color": "#203B58"}),
        "context": book.add_format({**base, "font_color": "#405C76"}),
        "note": book.add_format({**base, "font_color": "#63758A"}),
        "source": book.add_format({**base, "font_color": "#63758A", "bottom": 1, "bottom_color": "#CFDCE6"}),
        "link": book.add_format({**base, "font_color": "#1C6284", "underline": True}),
        "pending": book.add_format({"bg_color": "#FFF2D9", "font_color": "#8A591A"}),
        "error": book.add_format({"bg_color": "#FBE8E8", "font_color": "#A53036", "bold": True}),
    }
    styles["headers"] = [book.add_format({**base, "bg_color": color, "font_color": "#FFFFFF",
                                          "bold": True, "text_wrap": True, "align": "center",
                                          "right": 1, "right_color": "#FFFFFF"})
                         for color in ("#263E58", "#1D607C", "#52667B")]
    styles["rows"] = []
    for fill in ("#FFFFFF", "#F2F6FA"):
        row_base = {**base, "bg_color": fill}
        styles["rows"].append({
            "code": book.add_format({**row_base, "num_format": "@", "align": "center", "font_color": "#5D7185"}),
            "name": book.add_format({**row_base, "align": "left", "indent": 1, "text_wrap": True}),
            "money": book.add_format({**row_base, "num_format": "#,##0.00", "align": "right", "indent": 1,
                                       "font_size": 11, "bold": True, "font_color": "#194F6E"}),
            "count": book.add_format({**row_base, "num_format": "#,##0", "align": "right", "indent": 1}),
            "status": book.add_format({**row_base, "align": "center"}),
            "date": book.add_format({**row_base, "num_format": "yyyy-mm-dd", "align": "center", "font_color": "#405C76"}),
        })
    return styles


def _prepare_sheet(sheet, report, source, styles, headings, widths, money_columns, frozen, description, other_sheet):
    sheet.hide_gridlines(2)
    sheet.set_zoom(90)
    sheet.set_default_row(26)
    sheet.set_tab_color("#263E58" if frozen == 2 else "#1D607C")
    for column, width in enumerate(widths):
        sheet.set_column(column, column, width)
    sheet.set_row(0, 9)
    sheet.set_row(1, 32)
    sheet.write_string(1, 0, "逐影院票房" if frozen == 2 else "每家影院每日票房", styles["title"])
    sheet.write_url(1, len(widths) - 3, f"internal:'{other_sheet}'!A1", styles["link"], f"查看{other_sheet}")
    sheet.set_row(2, 24)
    inferred = "（日期范围按明细推定）" if report.inferred_period else ""
    sheet.write_string(2, 0, f"电影编码 {report.movie}    {report.start} 至 {report.end}    {report.days} 天    {len(report.cinemas):,} 家已知影院{inferred}", styles["context"])
    sheet.set_row(3, 24)
    sheet.write_string(3, 0, description, styles["note"])
    sheet.set_row(4, 24)
    # A single-row source line can span empty cells while the rule covers the table.
    for column in range(len(widths)):
        sheet.write_blank(4, column, None, styles["source"])
    sheet.write_string(4, 0, f"来源 {source.name}    全国名单覆盖待核实", styles["source"])
    sheet.set_row(HEADER_ROW, 36)
    for column, text in enumerate(headings):
        section = 1 if column in money_columns else 0 if column < frozen else 2
        sheet.write_string(HEADER_ROW, column, text, styles["headers"][section])
    sheet.freeze_panes(HEADER_ROW + 1, frozen)
    sheet.set_landscape()
    sheet.set_paper(8)  # A3 keeps long cinema names and the audit columns legible.
    sheet.fit_to_pages(1, 0)
    sheet.set_margins(0.3, 0.3, 0.4, 0.4)
    sheet.repeat_rows(HEADER_ROW)
    sheet.set_footer('&L&"Microsoft YaHei"&9MovieDataDecoder&R第 &P / &N 页')


def _name_row_height(name, width):
    lines = 0
    for line in name.splitlines():
        units = sum(2 if unicodedata.east_asian_width(char) in ("W", "F") else 1 for char in line)
        lines += max(1, math.ceil(units / (width - 4)))
    return min(409, max(26, lines * 16 + 8))


def _finish_sheet(sheet, last_row, columns, styles):
    sheet.autofilter(HEADER_ROW, 0, last_row, columns - 1)
    sheet.print_area(0, 0, last_row, columns - 1)
    # Conditions follow the visible status text after edits or sorting in Excel.
    for keyword, style in (("异常", "error"), ("不一致", "error"), ("未知", "error"), ("待", "pending")):
        sheet.conditional_format(HEADER_ROW + 1, 4, last_row, 4,
                                 {"type": "text", "criteria": "containing", "value": keyword,
                                  "format": styles[style]})


def _write_report(path, scratch, report, source, conn, emit):
    headings = ("影院编码", "影院名称", "已确认票房（元）", "完整区间票房（元）", "核查状态",
                "已核查天数", "待核查天数", "204 待复核天数", "有该片天数")
    widths = (14, 44, 21, 23, 27, 13, 13, 18, 13)
    ordered = sorted(report.cinemas.items(), key=lambda item: (not item[1].verified, -item[1].cents, item[0]))
    room = MAX_ROWS - HEADER_ROW - 1
    with xlsxwriter.Workbook(str(path), {"constant_memory": True, "tmpdir": str(scratch),
                                       "strings_to_formulas": False, "strings_to_urls": False}) as book:
        styles = _report_styles(book)
        for part, index in enumerate(range(0, len(ordered), room), 1):
            sheet = book.add_worksheet("逐影院票房" if part == 1 else f"逐影院票房 {part}")
            _prepare_sheet(sheet, report, source, styles, headings, widths, (2, 3), 2,
                           "金额单位为元。已确认票房按有效日期累计，核查齐全部日期才填完整区间票房。空白金额表示待核查。", "逐日明细")
            for row, (code, cinema) in enumerate(ordered[index:index + room], HEADER_ROW + 1):
                pending_days = report.days - cinema.verified
                if cinema.cents > MAX_CENTS:
                    raise ValueError(f"影院 {code} 的汇总金额超出可安全写入 Excel 的精度范围。")
                cell = styles["rows"][(row - HEADER_ROW - 1) % 2]
                sheet.set_row(row, _name_row_height(cinema.name, widths[1]))
                sheet.write_string(row, 0, code, cell["code"])
                sheet.write_string(row, 1, cinema.name[:32_767], cell["name"])
                sheet.write_blank(row, 2, None, cell["money"])
                sheet.write_blank(row, 3, None, cell["money"])
                if cinema.verified:
                    amount = Decimal(cinema.cents) / 100
                    sheet.write_number(row, 2, amount, cell["money"])
                    if not pending_days and not report.inferred_period:
                        sheet.write_number(row, 3, amount, cell["money"])
                if pending_days:
                    status = "部分日期待核查" if cinema.verified else "待核查"
                elif report.inferred_period:
                    status = "导出日期范围待核实"
                else:
                    status = "已核查" if cinema.with_movie else "已核查，无该片"
                sheet.write_string(row, 4, status, cell["status"])
                sheet.write_row(row, 5, (cinema.verified, pending_days, cinema.deferred, cinema.with_movie), cell["count"])
            _finish_sheet(sheet, row, len(headings), styles)
        _write_daily(book, report, source, conn, emit, styles)


def _write_daily(book, report, source, conn, emit, styles):
    headings = ("日期", "影院编码", "影院名称", "该片当日票房（元）", "采集状态", "原始状态", "电影编码")
    widths = (15, 14, 48, 25, 31, 21, 18)
    codes = sorted(report.cinemas)
    sheet, row, part = None, MAX_ROWS, 0
    total = len(codes) * report.days
    written = 0
    for offset in range(report.days):
        day = report.start + dt.timedelta(days=offset)
        cursor = conn.execute("SELECT cinema,status,cents,name FROM seen WHERE day=? ORDER BY cinema", (day.isoformat(),))
        record = next(cursor, None)
        for code in codes:
            if row == MAX_ROWS:
                if sheet is not None:
                    _finish_sheet(sheet, row - 1, len(headings), styles)
                part += 1
                sheet = book.add_worksheet("逐日明细" if part == 1 else f"逐日明细 {part}")
                _prepare_sheet(sheet, report, source, styles, headings, widths, (3,), 3,
                               "每家影院每天一行，金额单位为元。204、未采集及异常的金额留空，已核查无该片记 0。", "逐影院票房")
                row = HEADER_ROW + 1
            name = report.cinemas[code].name
            if record is not None and record[0] == code:
                _, status, cents, daily_name = record
                name = daily_name or name
                status_text = STATUS_LABELS.get(status, "未知状态，待复核")
                record = next(cursor, None)
            else:
                status, cents, status_text = "", None, "源文件缺行，待核查"
            cell = styles["rows"][(row - HEADER_ROW - 1) % 2]
            sheet.set_row(row, _name_row_height(name, widths[2]))
            sheet.write_datetime(row, 0, day, cell["date"])
            sheet.write_string(row, 1, code, cell["code"])
            sheet.write_string(row, 2, name[:32_767], cell["name"])
            sheet.write_blank(row, 3, None, cell["money"])
            if cents is not None:
                sheet.write_number(row, 3, Decimal(cents) / 100, cell["money"])
            sheet.write_string(row, 4, status_text, cell["status"])
            sheet.write_string(row, 5, status[:32_767], cell["code"])
            sheet.write_string(row, 6, report.movie, cell["code"])
            row += 1
            written += 1
        emit(f"已生成 {written:,} / {total:,} 行逐日明细…")
    _finish_sheet(sheet, row - 1, len(headings), styles)


def convert(source, output, *, emit=lambda message: None):
    output = Path(output).resolve()
    if output.suffix.lower() != ".xlsx" or not output.parent.is_dir():
        raise ValueError("请选择现有文件夹及 .xlsx 文件名。")
    with tempfile.TemporaryDirectory(prefix=".moviedata-decoder-", dir=output.parent) as folder:
        scratch = Path(folder)
        emit("正在读取影院日明细…")
        with closing(sqlite3.connect(scratch / "seen.db")) as conn:
            report = _read_report(conn, source, emit)
            emit(f"正在生成 {len(report.cinemas):,} 家影院的区间汇总与逐日明细…")
            local = scratch / "report.xlsx"
            _write_report(local, scratch, report, source, conn, emit)
        os.replace(local, output)
    return len(report.cinemas)


def main():
    parser = argparse.ArgumentParser(description="将 MovieData 的影院明细转换为逐影院票房 Excel")
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", "-o", type=Path)
    parser.add_argument("--list", action="store_true", help="列出可处理的电影明细")
    parser.add_argument("--select", type=int, default=1, help="选择列出的编号，默认第 1 项")
    args = parser.parse_args()
    try:
        items = discover(args.input)
        if not items:
            parser.error(NO_DETAILS)
        if args.list:
            for i, source in enumerate(items, 1):
                print(f"{i}. {source.location}")
            return
        if not 1 <= args.select <= len(items):
            parser.error("所选编号超出范围，可先使用 --list。")
        source = items[args.select - 1]
        output = args.output or Path.cwd() / source.filename
        count = convert(source, output, emit=print)
    except (OSError, ValueError, csv.Error) as exc:
        parser.exit(1, f"转换失败：{exc}\n")
    print(f"已生成 {count:,} 家影院的票房表，保存到 {output.resolve()}")


if __name__ == "__main__":
    main()
