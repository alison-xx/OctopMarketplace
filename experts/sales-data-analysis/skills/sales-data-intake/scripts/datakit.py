#!/usr/bin/env python3
"""Small, dependency-light helpers for business tables (CSV / TSV / XLSX).

Only the standard library is required; ``.xlsx`` / ``.xlsm`` files also need
``openpyxl``. Every command prints Markdown that can be pasted into a reply.

    python datakit.py profile  FILE [--sheet NAME]
    python datakit.py rfm      FILE --customer COL --date COL --amount COL
                               [--order COL] [--asof YYYY-MM-DD] [--out FILE.csv]
    python datakit.py variance FILE --period COL --base VALUE --current VALUE
                               --measure COL --by COL [--qty COL] [--top N]

Column arguments accept the header text exactly as it appears in the file.
"""

from __future__ import annotations

import argparse
import csv
import io
import math
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

# -- reading -------------------------------------------------------------------

_ENCODINGS = ("utf-8-sig", "gb18030", "utf-16")
_NUMBER_RE = re.compile(r"^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$")
_DATE_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%Y%m%d",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M",
    "%Y年%m月%d日",
)


class DataError(Exception):
    """A problem with the input the user can fix (missing column, bad file)."""


def _decode(raw: bytes) -> str:
    for encoding in _ENCODINGS:
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise DataError("cannot decode the file as UTF-8, GB18030 or UTF-16")


def _read_delimited(path: Path) -> list[list[str]]:
    text = _decode(path.read_bytes())
    sample = text[:20000]
    if path.suffix.lower() == ".tsv":
        delimiter = "\t"
    else:
        try:
            delimiter = csv.Sniffer().sniff(sample, delimiters=",\t;|").delimiter
        except csv.Error:
            delimiter = ","
    return [row for row in csv.reader(io.StringIO(text), delimiter=delimiter)]


def _read_xlsx(path: Path, sheet: str | None) -> list[list[object]]:
    try:
        import openpyxl  # type: ignore[import-untyped]
    except ImportError as exc:  # pragma: no cover - depends on the environment
        raise DataError("reading .xlsx needs openpyxl (pip install openpyxl)") from exc
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet:
            if sheet not in book.sheetnames:
                raise DataError(f"sheet {sheet!r} not found; sheets: {', '.join(book.sheetnames)}")
            ws = book[sheet]
        else:
            ws = book.worksheets[0]
        return [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        book.close()


def _blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _header_index(rows: list[list[object]]) -> int:
    """First row that looks like a header: mostly filled, mostly text.

    Exported reports often start with a title row ("2026年10月销售明细") and
    blank lines; those are skipped.
    """
    width = max((sum(1 for cell in row if not _blank(cell)) for row in rows[:50]), default=0)
    for index, row in enumerate(rows[:50]):
        filled = [cell for cell in row if not _blank(cell)]
        if width and len(filled) >= max(2, math.ceil(width * 0.6)):
            texty = sum(1 for cell in filled if isinstance(cell, str) and to_number(cell) is None)
            if texty >= len(filled) * 0.6:
                return index
    return 0


def read_table(path: str, sheet: str | None = None) -> tuple[list[str], list[dict[str, object]]]:
    """Return ``(headers, rows)``; blank rows are dropped, headers made unique."""
    file = Path(path)
    if not file.is_file():
        raise DataError(f"file not found: {path}")
    suffix = file.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        grid = _read_xlsx(file, sheet)
    elif suffix in {".csv", ".tsv", ".txt"}:
        grid = [list(row) for row in _read_delimited(file)]
    else:
        raise DataError(f"unsupported file type {suffix!r}; use .csv, .tsv or .xlsx")
    grid = [row for row in grid if any(not _blank(cell) for cell in row)]
    if not grid:
        raise DataError("the table is empty")
    start = _header_index(grid)
    headers: list[str] = []
    seen: Counter[str] = Counter()
    for position, cell in enumerate(grid[start]):
        name = str(cell).strip() if not _blank(cell) else f"column_{position + 1}"
        seen[name] += 1
        headers.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
    rows = []
    for raw in grid[start + 1 :]:
        cells = list(raw) + [None] * (len(headers) - len(raw))
        rows.append({name: cells[i] for i, name in enumerate(headers)})
    return headers, rows


def require(headers: list[str], *names: str | None) -> None:
    missing = [name for name in names if name and name not in headers]
    if missing:
        raise DataError(f"column not found: {', '.join(missing)}; columns: {', '.join(headers)}")


# -- values --------------------------------------------------------------------


def to_number(value: object) -> float | None:
    """Parse 1,234.5 / ￥1234 / 12% / (300) style values; None when not numeric."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return None if isinstance(value, float) and math.isnan(value) else float(value)
    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    for mark in ("￥", "¥", "$", "元", "RMB", "CNY"):
        text = text.replace(mark, "")
    percent = text.endswith("%")
    text = text.rstrip("%")
    if not _NUMBER_RE.match(text):
        return None
    number = float(text.replace(",", ""))
    if percent:
        number /= 100
    return -number if negative else number


def to_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel serial numbers that reached us as plain numbers.
        if 20000 < float(value) < 80000:
            return date(1899, 12, 30) + timedelta(days=int(value))
        return None
    text = str(value or "").strip()
    if not text:
        return None
    # "2026-10-08T09:30:00" and "2026-10-08 09:30:00.123" keep only the day part.
    candidates = [text, re.split(r"[ T]", text, maxsplit=1)[0]]
    for candidate in candidates:
        for pattern in _DATE_FORMATS:
            try:
                return datetime.strptime(candidate, pattern).date()
            except ValueError:
                continue
    return None


def label(value: object) -> str:
    if _blank(value):
        return "(空)"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, datetime):
        return value.date().isoformat()
    return str(value).strip()


def fmt(number: float, digits: int = 2) -> str:
    if number is None or (isinstance(number, float) and math.isnan(number)):
        return "-"
    if abs(number - round(number)) < 1e-9 and abs(number) < 1e15:
        return f"{int(round(number)):,}"
    return f"{number:,.{digits}f}"


def pct(part: float, whole: float) -> str:
    return "-" if not whole else f"{part / whole * 100:.1f}%"


def md_table(headers: list[str], rows: list[list[object]]) -> str:
    def cell(value: object) -> str:
        return str(value).replace("|", "\\|").replace("\n", " ")

    out = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
    out.extend("| " + " | ".join(cell(value) for value in row) + " |" for row in rows)
    return "\n".join(out)


# -- profile -------------------------------------------------------------------


def profile(path: str, sheet: str | None) -> str:
    headers, rows = read_table(path, sheet)
    lines = [f"# 数据画像：{Path(path).name}", "", f"行数：{len(rows):,}　列数：{len(headers)}"]
    keys = [tuple(label(row[name]) for name in headers) for row in rows]
    duplicates = len(keys) - len(set(keys))
    if duplicates:
        lines.append(f"完全重复的行：{duplicates:,}（核对是否为重复导出，不要直接删除）")
    lines.append("")
    table: list[list[object]] = []
    for name in headers:
        values = [row[name] for row in rows]
        present = [value for value in values if not _blank(value)]
        numbers = [n for n in (to_number(value) for value in present) if n is not None]
        dates = [d for d in (to_date(value) for value in present) if d is not None]
        distinct = len({label(value) for value in present})
        if present and len(dates) >= len(present) * 0.9 and len(numbers) < len(present) * 0.9:
            kind, detail = "日期", f"{min(dates).isoformat()} ~ {max(dates).isoformat()}"
        elif present and len(numbers) >= len(present) * 0.9:
            negatives = sum(1 for n in numbers if n < 0)
            kind = "数值"
            detail = (
                f"合计 {fmt(sum(numbers))}；最小 {fmt(min(numbers))}；最大 {fmt(max(numbers))}"
                + (f"；负数 {negatives} 个" if negatives else "")
            )
        else:
            kind = "文本"
            top = Counter(label(value) for value in present).most_common(3)
            detail = "；".join(f"{value}×{count}" for value, count in top)
        table.append([name, kind, len(values) - len(present), distinct, detail])
    lines.append(md_table(["列", "类型", "空值", "不同取值", "概况"], table))
    lines.append("")
    lines.append("类型由取值推断，金额是否含税、日期是下单还是发货等口径需要和用户确认。")
    return "\n".join(lines)


# -- rfm -----------------------------------------------------------------------

_SEGMENTS = {
    ("H", "H", "H"): "重要价值客户",
    ("L", "H", "H"): "重要保持客户",
    ("H", "L", "H"): "重要发展客户",
    ("L", "L", "H"): "重要挽留客户",
    ("H", "H", "L"): "一般价值客户",
    ("L", "H", "L"): "一般保持客户",
    ("H", "L", "L"): "一般发展客户",
    ("L", "L", "L"): "一般挽留客户",
}
_SEGMENT_ORDER = list(_SEGMENTS.values())


def rfm(
    path: str,
    *,
    sheet: str | None,
    customer: str,
    when: str,
    amount: str,
    order: str | None,
    asof: str | None,
    out: str | None,
) -> str:
    headers, rows = read_table(path, sheet)
    require(headers, customer, when, amount, order)
    last: dict[str, date] = {}
    orders: dict[str, set[str]] = defaultdict(set)
    money: dict[str, float] = defaultdict(float)
    skipped = 0
    for index, row in enumerate(rows):
        who = label(row[customer])
        day = to_date(row[when])
        value = to_number(row[amount])
        if who == "(空)" or day is None or value is None:
            skipped += 1
            continue
        last[who] = max(day, last.get(who, day))
        orders[who].add(label(row[order]) if order else f"{day.isoformat()}#{index}")
        money[who] += value
    if not last:
        raise DataError("no row has a customer, a date and a numeric amount")
    reference = to_date(asof) if asof else max(last.values())
    if reference is None:
        raise DataError(f"--asof is not a date: {asof}")
    customers = sorted(last)
    recency = {who: (reference - last[who]).days for who in customers}
    frequency = {who: len(orders[who]) for who in customers}
    mean_r = sum(recency.values()) / len(customers)
    mean_f = sum(frequency.values()) / len(customers)
    mean_m = sum(money.values()) / len(customers)
    segment = {
        who: _SEGMENTS[
            (
                "H" if recency[who] <= mean_r else "L",
                "H" if frequency[who] > mean_f else "L",
                "H" if money[who] > mean_m else "L",
            )
        ]
        for who in customers
    }
    total = sum(money.values())
    summary: list[list[object]] = []
    for name in _SEGMENT_ORDER:
        members = [who for who in customers if segment[who] == name]
        if not members:
            continue
        value = sum(money[who] for who in members)
        summary.append(
            [
                name,
                len(members),
                pct(len(members), len(customers)),
                fmt(value),
                pct(value, total),
                fmt(sum(recency[w] for w in members) / len(members), 1),
                fmt(sum(frequency[w] for w in members) / len(members), 1),
            ]
        )
    lines = [
        f"# RFM 客户分层：{Path(path).name}",
        "",
        f"客户数 {len(customers):,}；金额合计 {fmt(total)}；基准日 {reference.isoformat()}"
        + ("（数据中最后一笔交易日）" if not asof else ""),
        f"频次按{'订单号去重' if order else '交易行'}计；高低以全体客户均值为界"
        f"（R 均值 {fmt(mean_r, 1)} 天，F 均值 {fmt(mean_f, 1)}，M 均值 {fmt(mean_m)}）。",
    ]
    if skipped:
        lines.append(f"跳过 {skipped:,} 行：缺客户、日期或金额无法识别。")
    lines += [
        "",
        md_table(
            ["分层", "客户数", "客户占比", "金额", "金额占比", "平均未购天数", "平均频次"],
            summary,
        ),
    ]
    if out:
        with open(out, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            writer.writerow(["客户", "最近购买", "未购天数R", "频次F", "金额M", "分层"])
            for who in sorted(customers, key=lambda w: -money[w]):
                writer.writerow(
                    [who, last[who].isoformat(), recency[who], frequency[who], round(money[who], 2), segment[who]]
                )
        lines += ["", f"逐客户结果已写入 {out}"]
    return "\n".join(lines)


# -- variance ------------------------------------------------------------------


def variance(
    path: str,
    *,
    sheet: str | None,
    period: str,
    base: str,
    current: str,
    measure: str,
    by: str,
    qty: str | None,
    top: int,
) -> str:
    headers, rows = read_table(path, sheet)
    require(headers, period, measure, by, qty)
    value = {base: defaultdict(float), current: defaultdict(float)}
    volume = {base: defaultdict(float), current: defaultdict(float)}
    skipped = 0
    periods_seen: Counter[str] = Counter()
    for row in rows:
        tag = label(row[period])
        periods_seen[tag] += 1
        if tag not in value:
            continue
        amount = to_number(row[measure])
        units = to_number(row[qty]) if qty else 0.0
        if amount is None or (qty and units is None):
            skipped += 1
            continue
        member = label(row[by])
        value[tag][member] += amount
        volume[tag][member] += units or 0.0
    for tag in (base, current):
        if not value[tag]:
            seen = ", ".join(name for name, _ in periods_seen.most_common(12))
            raise DataError(f"period {tag!r} has no rows in column {period!r}; values seen: {seen}")
    total0, total1 = sum(value[base].values()), sum(value[current].values())
    change = total1 - total0
    members = set(value[base]) | set(value[current])
    ranked = sorted(members, key=lambda m: -abs(value[current][m] - value[base][m]))
    lines = [
        f"# {measure} 变动归因：{base} → {current}（按 {by}）",
        "",
        f"{base}：{fmt(total0)}　{current}：{fmt(total1)}　变动 {fmt(change)}（{pct(change, total0)}）",
    ]
    if skipped:
        lines.append(f"跳过 {skipped:,} 行：{measure}{'或' + qty if qty else ''}不是数字。")
    table = []
    for member in ranked[:top]:
        before, after = value[base][member], value[current][member]
        delta = after - before
        note = "新增" if before == 0 and after != 0 else "消失" if after == 0 and before != 0 else ""
        table.append([member, fmt(before), fmt(after), fmt(delta), pct(delta, before) if before else "-", pct(delta, change), note])
    rest = ranked[top:]
    if rest:
        delta = sum(value[current][m] - value[base][m] for m in rest)
        table.append([f"其余 {len(rest)} 项", "", "", fmt(delta), "", pct(delta, change), ""])
    lines += [
        "",
        md_table([by, base, current, "变动", "变动率", "占总变动", "备注"], table),
        "",
        "“占总变动”可正可负，合计为 100%；总变动接近 0 时比例没有意义，看绝对变动。",
    ]
    if qty:
        lines += ["", *_price_volume_mix(value, volume, base, current, members)]
    return "\n".join(lines)


def _price_volume_mix(
    value: dict[str, dict[str, float]],
    volume: dict[str, dict[str, float]],
    base: str,
    current: str,
    members: set[str],
) -> list[str]:
    """Split the change into volume, mix, price and new / lost members.

    Only members sold in both periods (with units in both) get a price; the
    revenue of the others is reported separately instead of being forced into
    a price or volume effect.
    """
    both = [m for m in members if volume[base][m] > 0 and volume[current][m] > 0]
    new = sum(value[current][m] for m in members if volume[base][m] <= 0)
    lost = -sum(value[base][m] for m in members if volume[current][m] <= 0)
    q0 = sum(volume[base][m] for m in both)
    q1 = sum(volume[current][m] for m in both)
    r0 = sum(value[base][m] for m in both)
    if not both or q0 == 0:
        return ["没有两期都有销量的项目，无法拆分价格与销量影响。"]
    price0 = {m: value[base][m] / volume[base][m] for m in both}
    price1 = {m: value[current][m] / volume[current][m] for m in both}
    average0 = r0 / q0
    volume_effect = (q1 - q0) * average0
    mix_effect = sum(volume[current][m] * price0[m] for m in both) - q1 * average0
    price_effect = sum((price1[m] - price0[m]) * volume[current][m] for m in both)
    total = volume_effect + mix_effect + price_effect + new + lost
    rows = [
        ["销量（总量变化）", fmt(volume_effect)],
        ["结构（高低价产品占比变化）", fmt(mix_effect)],
        ["价格（同一项目单价变化）", fmt(price_effect)],
        ["新增项目", fmt(new)],
        ["停售/流失项目", fmt(lost)],
        ["合计", fmt(total)],
    ]
    return [
        "## 量价结构拆分",
        "",
        md_table(["因素", "影响金额"], rows),
        "",
        "单价＝金额÷数量；价格影响按本期数量计，结构影响按上期单价计。"
        "数量单位必须一致（件/箱混用时先换算）。",
    ]


# -- cli -----------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("profile", help="columns, types, blanks and ranges of a table")
    p.add_argument("file")
    p.add_argument("--sheet")

    r = sub.add_parser("rfm", help="RFM customer segments")
    r.add_argument("file")
    r.add_argument("--sheet")
    r.add_argument("--customer", required=True)
    r.add_argument("--date", required=True)
    r.add_argument("--amount", required=True)
    r.add_argument("--order")
    r.add_argument("--asof")
    r.add_argument("--out")

    v = sub.add_parser("variance", help="explain the change of a measure between two periods")
    v.add_argument("file")
    v.add_argument("--sheet")
    v.add_argument("--period", required=True)
    v.add_argument("--base", required=True)
    v.add_argument("--current", required=True)
    v.add_argument("--measure", required=True)
    v.add_argument("--by", required=True)
    v.add_argument("--qty")
    v.add_argument("--top", type=int, default=10)

    args = parser.parse_args(argv)
    try:
        if args.command == "profile":
            print(profile(args.file, args.sheet))
        elif args.command == "rfm":
            print(
                rfm(
                    args.file,
                    sheet=args.sheet,
                    customer=args.customer,
                    when=args.date,
                    amount=args.amount,
                    order=args.order,
                    asof=args.asof,
                    out=args.out,
                )
            )
        else:
            print(
                variance(
                    args.file,
                    sheet=args.sheet,
                    period=args.period,
                    base=args.base,
                    current=args.current,
                    measure=args.measure,
                    by=args.by,
                    qty=args.qty,
                    top=max(args.top, 1),
                )
            )
    except DataError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
