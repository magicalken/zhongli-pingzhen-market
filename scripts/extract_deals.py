#!/usr/bin/env python3
"""Extract a month of transactions into the website JSON format.

The source is either the monthly workbook, or a tab/comma separated transcript
with the same A-H columns — used when the month only arrives as photos.

Usage:
    extract_deals.py SOURCE.xlsx data/2026-07.json
    extract_deals.py SOURCE.tsv  data/2026-07.json --register "2026 年 7 月"

With --register the month is also added to (or updated in) data/months.json,
so the site picks it up without any further hand editing.
"""
import argparse
import csv
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
}
T = f"{{{NS['x']}}}t"
RPH = f"{{{NS['x']}}}rPh"

# Header/title/sample rows that sit above or below the real data.
SKIP_STORE = re.compile(r"^(?:\d{2,3}\s*年|範例|店家|店名|合計|總計)")
EXCLUSIVE_MARKS = {"V", "Ｖ", "✓", "✔", "ˇ", "Y", "O", "是", "有"}


def rich_text(element):
    """Text of an <si> or <is>: direct <t>, plus <r> runs, skipping <rPh> furigana."""
    parts = []
    for child in element:
        if child.tag == RPH:
            continue
        if child.tag == T:
            parts.append(child.text or "")
        else:
            for node in child.iter(T):
                parts.append(node.text or "")
    return "".join(parts)


def load_shared_strings(archive):
    try:
        raw = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    return [rich_text(si) for si in root.findall("x:si", NS)]


def first_sheet_path(archive):
    """Resolve the workbook's first worksheet part, not just xl/worksheets/sheet1.xml."""
    try:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    except KeyError:
        return "xl/worksheets/sheet1.xml"

    sheet = workbook.find("x:sheets/x:sheet", NS)
    if sheet is None:
        return "xl/worksheets/sheet1.xml"
    rid = sheet.attrib.get(f"{{{NS['r']}}}id")
    for rel in rels:
        if rel.attrib.get("Id") == rid:
            target = rel.attrib["Target"].lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"
    return "xl/worksheets/sheet1.xml"


def cell_value(cell, shared):
    """Value of one cell, resolving shared strings, inline strings and formula results."""
    kind = cell.attrib.get("t", "n")

    if kind == "inlineStr":
        inline = cell.find("x:is", NS)
        return rich_text(inline) if inline is not None else ""

    value = cell.find("x:v", NS)
    if value is None or value.text is None:
        return None
    raw = value.text

    if kind == "s":
        index = int(raw)
        return shared[index] if 0 <= index < len(shared) else ""
    if kind in ("str", "e"):
        return raw
    if kind == "b":
        return raw == "1"

    return number(raw, fallback=raw)


def number(value, fallback=None):
    """Coerce a cell value to a number; Excel may hand back numerals as text."""
    if value is None or isinstance(value, bool):
        return fallback
    if isinstance(value, (int, float)):
        return value
    cleaned = re.sub(r"[$＄,，\s元萬]", "", str(value))
    if not cleaned:
        return fallback
    try:
        return int(cleaned)
    except ValueError:
        pass
    try:
        return float(cleaned)
    except ValueError:
        return fallback


def as_text(value):
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


COLUMNS = "ABCDEFGH"


def rows_to_records(rows):
    """Turn A-H rows of raw cell values into website records."""
    records, skipped, current_store = [], 0, ""
    for row in rows:
        cells = list(row) + [None] * (len(COLUMNS) - len(row))

        store = as_text(cells[0])
        if store and SKIP_STORE.match(store):
            continue
        if store:
            current_store = store

        sold = number(cells[5])
        if sold is None:
            continue
        # The sheet merges the store column, so a deal row may leave it blank.
        if not current_store:
            skipped += 1
            continue
        store = current_store

        records.append({
            "store": store,
            "name": as_text(cells[1]),
            "area": as_text(cells[2]),
            "size": number(cells[3]),
            "asking": number(cells[4]),
            "sold": sold,
            "unit": number(cells[6]),
            "exclusive": as_text(cells[7]).upper() in EXCLUSIVE_MARKS,
        })

    return records, skipped


def xlsx_rows(source):
    with zipfile.ZipFile(source) as archive:
        shared = load_shared_strings(archive)
        root = ET.fromstring(archive.read(first_sheet_path(archive)))

    for row in root.findall(".//x:sheetData/x:row", NS):
        cells = {}
        for cell in row.findall("x:c", NS):
            column = re.match(r"[A-Z]+", cell.attrib["r"]).group(0)
            cells[column] = cell_value(cell, shared)
        yield [cells.get(column) for column in COLUMNS]


def delimited_rows(source):
    """Rows from a tab/comma separated transcript of the monthly sheet."""
    text = Path(source).read_text(encoding="utf-8-sig")
    delimiter = "\t" if text.count("\t") >= text.count(",") else ","
    yield from csv.reader(text.splitlines(), delimiter=delimiter)


def extract(source):
    rows = xlsx_rows(source) if str(source).lower().endswith(".xlsx") else delimited_rows(source)
    return rows_to_records(rows)


def register_month(months_file, month_id, label, data_file):
    months = json.loads(months_file.read_text(encoding="utf-8")) if months_file.exists() else []
    entry = {"id": month_id, "label": label, "file": data_file}
    months = [m for m in months if m.get("id") != month_id] + [entry]
    months.sort(key=lambda m: m["id"], reverse=True)
    months_file.write_text(json.dumps(months, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return months


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("source", help="monthly workbook (.xlsx) or a tab/comma separated transcript (.tsv/.csv)")
    parser.add_argument("destination", help="output JSON, e.g. data/2026-07.json")
    parser.add_argument("--register", metavar="LABEL", help="also add this month to data/months.json, e.g. \"2026 年 7 月\"")
    args = parser.parse_args()

    records, skipped = extract(args.source)
    if not records:
        raise SystemExit(f"No transactions found in {args.source} — check that column A holds the store name.")

    destination = Path(args.destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(records, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    exclusive = sum(1 for r in records if r["exclusive"])
    priced = [r for r in records if r["asking"] and r["sold"]]
    discount = sum((1 - r["sold"] / r["asking"]) * 100 for r in priced) / len(priced) if priced else 0
    print(f"Extracted {len(records)} records to {destination}")
    print(f"  skipped {skipped} row(s) with a 成交價 but no 店家")
    print(f"  專約 {exclusive} 筆（{exclusive / len(records) * 100:.1f}%）· 平均議價率 {discount:.1f}%（{len(priced)} 筆有開價）")
    if exclusive == 0:
        print("  ⚠ 沒有任何一筆被判定為專約，請確認 H 欄的標記")

    if args.register:
        month_id = destination.stem
        months = register_month(destination.parent / "months.json", month_id, args.register, str(destination).replace("\\", "/"))
        print(f"  months.json 已更新，目前 {len(months)} 個月份，預設顯示 {months[0]['label']}")


if __name__ == "__main__":
    main()
