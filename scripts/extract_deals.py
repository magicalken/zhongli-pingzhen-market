#!/usr/bin/env python3
"""Extract the uploaded monthly transaction workbook into the website JSON format."""
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def cell_value(cell):
    inline = cell.find("x:is/x:t", NS)
    if inline is not None:
        return inline.text or ""
    value = cell.find("x:v", NS)
    if value is None or value.text is None:
        return None
    raw = value.text
    return float(raw) if "." in raw else int(raw)


def main(source, destination):
    with zipfile.ZipFile(source) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))

    records = []
    for row in root.findall(".//x:sheetData/x:row", NS):
        cells = {}
        for cell in row.findall("x:c", NS):
            column = re.match(r"[A-Z]+", cell.attrib["r"]).group(0)
            cells[column] = cell_value(cell)

        store = str(cells.get("A") or "").strip()
        if not store or store.startswith("115年") or store.startswith("範例"):
            continue

        record = {
            "store": store,
            "name": str(cells.get("B") or "").strip(),
            "area": str(cells.get("C") or "").strip(),
            "size": cells.get("D"),
            "asking": cells.get("E"),
            "sold": cells.get("F"),
            "unit": cells.get("G"),
            "exclusive": str(cells.get("H") or "").strip().upper() == "V",
        }
        if record["sold"] is not None:
            records.append(record)

    Path(destination).parent.mkdir(parents=True, exist_ok=True)
    Path(destination).write_text(json.dumps(records, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Extracted {len(records)} records to {destination}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("Usage: extract_deals.py SOURCE.xlsx DESTINATION.json")
    main(sys.argv[1], sys.argv[2])
