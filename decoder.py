"""Turn MovieData result files into readable, offline Excel snapshots."""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import os
import re
import sqlite3
import tempfile
from collections import Counter, defaultdict
from contextlib import closing, contextmanager
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

import xlsxwriter


VERSION = "1.0.0"
STATES = {
    "with_movie": "有目标电影", "without_movie": "无目标电影",
    "empty_204": "204 待复核", "deferred_204": "204 待复核", "missing": "未采集",
    "error": "请求异常", "inconsistent": "数据不一致",
}
VERIFIED = {"with_movie", "without_movie"}
UNKNOWN = "other"
MAX_ROWS = 1_048_576
DETAIL_SUFFIX = "_cinema_days.csv"
AUDIT_SUFFIX = "_daily_audit.csv"
SUMMARY_SUFFIX = "_summary.json"
DETAIL_FIELDS = {"date", "cinema_code", "movie_code", "status", "sales_fen"}
AUDIT_FIELDS = {"date", "target_count", *STATES}
LABELS = {
    "date": "日期", "show_date": "日期", "cinema_code": "影院编码", "cinema_name": "影院名称",
    "movie_code": "电影编码", "movie_name": "电影名称", "sales_fen": "票房（分）", "sales_yuan": "票房（元）",
    "status": "状态", "status_zh": "状态说明", "source": "来源", "source_url": "来源网址",
    "national_roster_coverage": "全国名单覆盖", "national_complete": "全国完整", "province": "省份",
    "province_code": "省份代码", "city": "城市", "city_code": "城市编码", "business_status": "营业状态",
    "roster_sources": "候选来源", "invalid_source_rows": "源站无效行数", "duplicate_source_rows": "源站重复行数",
    "target_count": "已知候选影院日", "observed": "已核查影院日", "tickets": "人次", "sessions": "场次",
}


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

    @contextmanager
    def open(self):
        if self.member:
            with ZipFile(self.path) as archive, archive.open(self.member) as stream:
                with io.TextIOWrapper(stream, encoding="utf-8-sig", newline="") as handle:
                    yield handle
        else:
            with self.path.open(encoding="utf-8-sig", newline="") as handle:
                yield handle

    def sibling(self, name):
        if self.member:
            return Source(self.path, str(PurePosixPath(self.member).with_name(name)))
        return Source(self.path.with_name(name))


@dataclass(frozen=True)
class Dataset:
    source: Source
    kind: str
    audit: Source | None = None

    @property
    def description(self):
        prefix = {"detail": "影院明细", "audit": "仅逐日审计", "summary": "仅逐日汇总 JSON", "csv": "普通表格"}[self.kind]
        return f"{prefix}  |  {self.source.location}"

    @property
    def filename(self):
        stem = Path(self.source.name).stem
        stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem).strip(" .")[:105] or "MovieData"
        return f"{stem}_易读版.xlsx"


def discover(path):
    path = Path(path).resolve()
    if path.is_dir():
        folders = (path, path / "national_history", path / "output", path / "output" / "national_history",
                   path / "tables", path / "output" / "tables")
        files = {Source(p) for folder in folders if folder.is_dir() for p in folder.iterdir() if p.is_file()}
        selection = files
    elif path.suffix.lower() == ".zip":
        with ZipFile(path) as archive:
            names = [info.filename for info in archive.infolist() if not info.is_dir()]
        if len(names) != len(set(names)):
            raise ValueError("压缩包中存在同名文件，请解压后单独选择要转换的结果。")
        files = {Source(path, name) for name in names}
        selection = files
    elif path.is_file():
        selection = {Source(path)}
        files = {Source(p) for p in path.parent.iterdir() if p.is_file()}
    else:
        raise ValueError("找不到所选文件或文件夹。")

    items = []
    for file in selection:
        name = file.name.lower()
        if name.endswith(DETAIL_SUFFIX):
            audit = file.sibling(file.name[:-len(DETAIL_SUFFIX)] + AUDIT_SUFFIX)
            items.append(Dataset(file, "detail", audit if audit in files else None))
        elif name.endswith(AUDIT_SUFFIX):
            items.append(Dataset(file, "audit"))
        elif name.endswith(SUMMARY_SUFFIX) and re.match(r"[0-9]+_\d{4}-", name):
            items.append(Dataset(file, "summary"))
        elif name.endswith(".csv"):
            items.append(Dataset(file, "csv"))
    order = {"detail": 0, "audit": 1, "summary": 2, "csv": 3}
    return sorted(items, key=lambda item: (order[item.kind], str(item.source.path), item.source.location))


def _headers(reader):
    values = [value.strip() for value in next(reader, [])]
    if not values or "" in values or len(set(values)) != len(values):
        raise ValueError("CSV 缺少表头，或表头含空列、重复列。")
    if len(values) > 16_384:
        raise ValueError("CSV 超过 Excel 单表可容纳的列数。")
    return values


def _csv_rows(source):
    with source.open() as handle:
        reader = csv.reader(handle, strict=True)
        fields = _headers(reader)
        for values in reader:
            if not values:
                continue
            if len(values) != len(fields):
                raise ValueError(f"{source.name} 第 {reader.line_num:,} 行与表头列数不同。")
            yield reader.line_num, dict(zip(fields, values))


def _day(value):
    day = dt.date.fromisoformat(value)
    if day.isoformat() != value:
        raise ValueError("日期必须为 YYYY-MM-DD。")
    return day


def _number(value):
    if not re.fullmatch(r"[0-9]+", str(value)):
        raise ValueError(f"需要非负整数，读到 {str(value)[:40]!r}。")
    number = int(value)
    if number > 999_999_999_999_999:
        raise ValueError("数值超出可安全写入 Excel 的范围。")
    return number


def _audit(source, json_mode=False):
    if json_mode:
        with source.open() as handle:
            payload = json.load(handle)
        if not isinstance(payload, dict) or not isinstance(payload.get("days"), list):
            raise ValueError("此文件不包含 MovieData 的逐日审计汇总。")
        entries = enumerate(payload["days"], 1)
    else:
        entries = _csv_rows(source)
    days = {}
    for line, row in entries:
        try:
            if not isinstance(row, dict) or not AUDIT_FIELDS <= row.keys():
                raise ValueError("缺少逐日审计字段。")
            day = _day(row["date"])
            counts = Counter({state: _number(row[state]) for state in STATES})
            observed = sum(counts[state] for state in VERIFIED)
            if day in days or _number(row["target_count"]) != sum(counts.values()):
                raise ValueError("日期重复，或状态合计与候选数不一致。")
            if "observed" in row and observed != _number(row["observed"]):
                raise ValueError("已核查数与状态合计不一致。")
            days[day] = counts
        except (ValueError, TypeError) as exc:
            raise ValueError(f"{source.name} 第 {line:,} 条审计记录：{exc}") from exc
    if not days:
        raise ValueError("审计文件没有日期记录。")
    return days


def _stage(conn, source, emit):
    conn.execute("""CREATE TABLE details (day TEXT, code TEXT, name TEXT, status TEXT, cents INTEGER,
                 PRIMARY KEY(day,code)) WITHOUT ROWID""")
    counts = defaultdict(Counter)
    cents_by_day = Counter()
    movie = ""
    repeats = discarded_money = input_rows = 0
    for line, row in _csv_rows(source):
        try:
            if not DETAIL_FIELDS <= row.keys():
                raise ValueError("缺少影院日明细字段。")
            day = _day(row["date"].strip())
            code, film, state = (row[key].strip() for key in ("cinema_code", "movie_code", "status"))
            if not re.fullmatch(r"[0-9]{6,10}", code) or not re.fullmatch(r"[0-9]{1,30}", film) or not state:
                raise ValueError("影院编码、电影编码或状态无效。")
            if movie and film != movie:
                raise ValueError("一个影院日明细文件应对应一部电影。")
            movie = film
            value = row["sales_fen"].strip()
            if state == "with_movie":
                cents = _number(value)
            elif state == "without_movie":
                if value and _number(value) != 0:
                    raise ValueError("无目标电影的行包含非零票房。")
                cents = 0
            else:
                cents = None
            result = conn.execute("INSERT OR IGNORE INTO details VALUES (?,?,?,?,?)",
                                  (day.isoformat(), code, row.get("cinema_name", "").strip(), state, cents))
            input_rows += 1
            if result.rowcount == 0:
                previous = conn.execute("SELECT status,cents FROM details WHERE day=? AND code=?",
                                        (day.isoformat(), code)).fetchone()
                if previous != (state, cents):
                    raise ValueError(f"影院 {code} 在 {day} 有冲突记录，请核对后重新导出。")
                repeats += 1
            else:
                counts[day][state if state in STATES else UNKNOWN] += 1
                if cents is not None:
                    cents_by_day[day] += cents
                elif value:
                    discarded_money += 1
            if input_rows % 25_000 == 0:
                emit(f"已读取 {input_rows:,} 条影院日记录…")
        except ValueError as exc:
            raise ValueError(f"{source.name} 第 {line:,} 行：{exc}") from exc
    if not input_rows:
        raise ValueError("影院明细为空。可改选逐日审计文件查看候选数。")
    conn.commit()
    return movie, dict(counts), cents_by_day, input_rows, repeats, discarded_money


class Book:
    def __init__(self, path, tmpdir):
        self.book = xlsxwriter.Workbook(str(path), {"constant_memory": True, "tmpdir": str(tmpdir),
                                                   "strings_to_formulas": False, "strings_to_urls": False})
        style = self.book.add_format
        self.title = style({"font_name": "Microsoft YaHei", "font_size": 16, "bold": True, "font_color": "#17365D"})
        self.hint = style({"font_color": "#56677C", "font_size": 10})
        self.header = style({"bold": True, "font_color": "#FFFFFF", "bg_color": "#214D7B", "text_wrap": True,
                             "valign": "vcenter"})
        self.money = style({"num_format": '#,##0.00;[Red](#,##0.00);0.00', "font_color": "#17365D"})
        self.date = style({"num_format": "yyyy-mm-dd"})
        self.integer = style({"num_format": "#,##0"})
        self.ratio = style({"num_format": "0.0%"})
        self.open_state = style({"font_color": "#943C21", "bg_color": "#FFF3E8"})

    def sheet(self, name, hint, columns, widths, *, freeze=0):
        sheet = self.book.add_worksheet(name)
        sheet.hide_gridlines(2)
        sheet.set_column(0, len(columns) - 1, 17)
        for col, width in enumerate(widths):
            sheet.set_column(col, col, width)
        sheet.set_row(0, 28)
        sheet.write_string(0, 0, name, self.title)
        sheet.write_string(1, 0, hint, self.hint)
        sheet.set_row(3, 30)
        sheet.write_row(3, 0, columns, self.header)
        sheet.freeze_panes(4, freeze)
        return sheet

    def overview(self, source, items):
        sheet = self.sheet("结果概览", f"来源文件  {source.location}", ("项目", "结果 / 说明"), (32, 105))
        for row, (label, value) in enumerate(items, 4):
            fmt = self.money if isinstance(value, Decimal) else self.integer if isinstance(value, int) else None
            sheet.write(row, 0, label)
            sheet.write(row, 1, value, fmt)

    def daily(self, days, cents=None):
        rate_col = 9
        columns = ("日期", "已知候选影院日", "已核查", "有目标电影", "无目标电影", "204 待复核", "未采集",
                   "异常 / 不一致", "其他未识别", "已知范围核查率", "已核查票房（元）")
        hint = ("金额仅汇总当前明细的有效片目。没有有效片目的日期留空。" if cents is not None
                else "数据来自逐日审计。原文件没有影院金额，金额列留空。")
        sheet = self.sheet("每日汇总", hint, columns, (16, 21, 15, 17, 17, 17, 15, 19, 16, 22, 25), freeze=1)
        for row, (day, counts) in enumerate(sorted(days.items()), 4):
            total = sum(counts.values())
            observed = sum(counts[s] for s in VERIFIED)
            sheet.write_datetime(row, 0, day, self.date)
            sheet.write_row(row, 1, (total, observed, counts["with_movie"], counts["without_movie"],
                            counts["empty_204"] + counts["deferred_204"], counts["missing"],
                            counts["error"] + counts["inconsistent"], counts[UNKNOWN]), self.integer)
            if total:
                sheet.write_number(row, rate_col, observed / total, self.ratio)
            if cents is not None and observed:
                sheet.write_number(row, 10, Decimal(cents.get(day, 0)) / 100, self.money)
        sheet.autofilter(3, 0, 3 + len(days), len(columns) - 1)

    def details(self, conn, only_movie, emit):
        cols = ("日期", "影院编码", "影院名称", "核查结果", "票房（元）", "源状态")
        widths = (16, 17, 48, 22, 19, 24)
        hint = "按日期和影院编码排序。未确知的金额留空。"
        if only_movie:
            hint = "此表只展示含目标电影的记录。每日汇总仍包含全部状态。"
        where = " WHERE status='with_movie'" if only_movie else ""
        rows = conn.execute("SELECT day,code,name,status,cents FROM details" + where + " ORDER BY day,code")
        sheet = self.sheet("影院明细", hint, cols, widths, freeze=2)
        logged = part = 0
        row = 4
        for day, code, name, state, cents in rows:
            if row == MAX_ROWS:
                sheet.autofilter(3, 0, row - 1, len(cols) - 1)
                part += 1
                sheet = self.sheet(f"影院明细 {part + 1}", hint, cols, widths, freeze=2)
                row = 4
            sheet.write_datetime(row, 0, _day(day), self.date)
            sheet.write_string(row, 1, code)
            if name:
                sheet.write_string(row, 2, name[:32_767])
            sheet.write_string(row, 3, STATES.get(state, "未识别状态"), None if state in VERIFIED else self.open_state)
            if cents is not None:
                sheet.write_number(row, 4, Decimal(cents) / 100, self.money)
            sheet.write_string(row, 5, state)
            row += 1
            logged += 1
            if logged % 25_000 == 0:
                emit(f"已写入 {logged:,} 条影院明细…")
        sheet.autofilter(3, 0, row - 1, len(cols) - 1)
        return logged

    def comparison(self, matrix, audit):
        cols = ("日期", "明细候选数", "审计候选数", "明细已核查", "审计已核查", "对照结果")
        sheet = self.sheet("审计对照", "明细与逐日审计分别取数。差异通常需要检查两份导出的时间。", cols,
                           (16, 20, 20, 19, 19, 27), freeze=1)
        for row, day in enumerate(sorted(matrix.keys() | audit.keys()), 4):
            a, b = matrix.get(day), audit.get(day)
            sheet.write_datetime(row, 0, day, self.date)
            for index, counts in enumerate((a, b)):
                if counts is not None:
                    sheet.write_number(row, 1 + index, sum(counts.values()), self.integer)
                    sheet.write_number(row, 3 + index, sum(counts[s] for s in VERIFIED), self.integer)
            sheet.write(row, 5, "一致" if a == b else "缺少对应日期" if a is None or b is None else "状态计数不一致",
                        None if a == b else self.open_state)
        sheet.autofilter(3, 0, 3 + len(matrix.keys() | audit.keys()), len(cols) - 1)

    def generic(self, source):
        with source.open() as handle:
            reader = csv.reader(handle, strict=True)
            fields = _headers(reader)
            columns = [LABELS.get(field, field) for field in fields]
            widths = [max(16, min(40, 2 * len(col) + 4)) for col in columns]
            for index, field in enumerate(fields):
                if field in ("cinema_name", "movie_name", "source_url", "影院名", "影片名", "影院名称", "影片名称", "片名"):
                    widths[index] = 42
                if field in ("date", "show_date", "日期") or field.endswith("_code"):
                    widths[index] = 17
            sheet = self.sheet("数据表", "从 CSV 读取。标识符、未识别字段及空值保留原义。", columns, widths)
            row, part, total = 4, 1, 0
            for values in reader:
                if not values:
                    continue
                if len(values) != len(fields):
                    raise ValueError(f"{source.name} 第 {reader.line_num:,} 行与表头列数不同。")
                if row == MAX_ROWS:
                    sheet.autofilter(3, 0, row - 1, len(fields) - 1)
                    part += 1
                    sheet = self.sheet(f"数据表 {part}", "源 CSV 续表。", columns, widths)
                    row = 4
                for col, (field, value) in enumerate(zip(fields, values)):
                    if value:
                        self.generic_cell(sheet, row, col, field, value)
                row += 1
                total += 1
            sheet.autofilter(3, 0, row - 1, len(fields) - 1)
            return total

    def generic_cell(self, sheet, row, col, field, value):
        if field in ("date", "show_date", "日期"):
            try:
                sheet.write_datetime(row, col, _day(value), self.date)
                return
            except ValueError:
                pass
        if field in ("status", "national_roster_coverage", "national_complete"):
            value = STATES.get(value, {"unverified": "未核实", "false": "未证实", "true": "是"}.get(value, value))
        if (field in ("sales_yuan", "sales_fen") or re.search(r"票房.*[（(]万?元[）)]", field)
                or field in ("tickets", "sessions", "人次", "场次")):
            try:
                number = Decimal(value.replace(",", ""))
                if number.is_finite() and number.copy_abs() <= 999_999_999_999_999 and len(number.as_tuple().digits) <= 15:
                    sheet.write_number(row, col, number, self.money if "元" in field or field == "sales_yuan" else self.integer)
                    return
            except InvalidOperation:
                pass
        sheet.write_string(row, col, value[:32_767])


def convert(dataset, output, *, only_movie=False, emit=lambda message: None):
    output = Path(output).resolve()
    if output.suffix.lower() != ".xlsx":
        raise ValueError("请选择 .xlsx 输出文件名。")
    if not output.parent.is_dir():
        raise ValueError("输出文件夹不存在。")
    with tempfile.TemporaryDirectory(prefix=".moviedata-decoder-", dir=output.parent) as scratch:
        temp = Path(scratch)
        local = temp / "report.xlsx"
        emit("正在读取导出结果…")
        kind = dataset.kind
        if kind == "csv":
            with dataset.source.open() as handle:
                headers = set(_headers(csv.reader(handle, strict=True)))
            if DETAIL_FIELDS <= headers:
                kind = "detail"
            elif AUDIT_FIELDS <= headers:
                kind = "audit"
        with closing(sqlite3.connect(temp / "staging.db")) as conn:
            result_rows = _write_report(Book(local, temp), conn, dataset, kind, only_movie, emit)
        os.replace(local, output)
    return result_rows


def _write_report(book, conn, dataset, kind, only_movie, emit):
    try:
        if kind == "csv":
            return book.generic(dataset.source)
        if kind in ("audit", "summary"):
            counts = _audit(dataset.source, json_mode=kind == "summary")
            book.overview(dataset.source, (("模式", "逐日审计"), ("审计日期数", len(counts)),
                          ("数据范围", "已知影院，按所选文件的日审计统计"),
                          ("金额", "逐日审计没有影院日金额。请在 MovieData 手动导出明细后转换。")))
            book.daily(counts)
            return len(counts)
        movie, counts, cents, rows, repeats, ignored = _stage(conn, dataset.source, emit)
        audit, audit_error = None, ""
        if dataset.audit:
            try:
                audit = _audit(dataset.audit)
            except (ValueError, OSError, csv.Error) as exc:
                audit_error = f"配套审计无法读取：{exc}"
        checks = sum(sum(day[s] for s in VERIFIED) for day in counts.values())
        diff = sum(counts.get(day) != audit.get(day) for day in counts.keys() | audit.keys()) if audit else 0
        book.overview(dataset.source, (
            ("电影编码", movie), ("明细日期范围", f"{min(counts)} 至 {max(counts)}"),
            ("已知候选影院日", rows - repeats), ("已核查影院日", checks),
            ("已核查票房（元）", Decimal(sum(cents.values())) / 100 if checks else None),
            ("待核查影院日", rows - repeats - checks), ("合并相同影院日记录", repeats),
            ("待核查行中未计入的金额", ignored),
            ("数据范围", "金额是已核查片目的合计，全国覆盖情况仍需核实。"),
            ("空白与零", "204、未采、异常及未识别状态的金额留空。有效片目无该片时为零。"),
            ("配套逐日审计", audit_error or (f"{dataset.audit.name}，{diff} 个日期有差异" if audit else "未选择配套审计")),
        ))
        book.daily(counts, cents)
        if audit:
            book.comparison(counts, audit)
        emit("正在生成影院明细工作表…")
        book.details(conn, only_movie, emit)
        return rows - repeats
    finally:
        book.book.close()


def main():
    parser = argparse.ArgumentParser(description="把 MovieData 导出结果转换为 Excel")
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", "-o", type=Path)
    parser.add_argument("--list", action="store_true", help="列出可处理的结果")
    parser.add_argument("--select", type=int, default=1, help="选择列出的编号，默认第 1 项")
    parser.add_argument("--only-movie", action="store_true", help="影院明细仅显示含目标电影的影院日")
    args = parser.parse_args()
    items = discover(args.input)
    if not items:
        parser.error("没有找到可用的 CSV 或逐日审计汇总。")
    if args.list:
        for i, item in enumerate(items, 1):
            print(f"{i}. {item.description}")
        return
    if not 1 <= args.select <= len(items):
        parser.error("所选编号超出范围，可先使用 --list。")
    item = items[args.select - 1]
    dest = args.output or Path.cwd() / item.filename
    try:
        count = convert(item, dest, only_movie=args.only_movie, emit=print)
    except (OSError, ValueError, csv.Error) as exc:
        parser.exit(1, f"转换失败：{exc}\n")
    print(f"已处理 {count:,} 条记录，保存到 {dest.resolve()}")


if __name__ == "__main__":
    main()
