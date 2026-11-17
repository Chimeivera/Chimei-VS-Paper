#!/usr/bin/env python3
"""Publish an approved complete Excel snapshot into the static dashboard.

No network access is needed. Fail closed: never remove or silently modify a
previously published application, and never publish a duplicate identifier.
"""
import json
import re
import sys
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
INDEX = ROOT / "index.html"
SOURCE = ROOT / "data" / "論文申請資料.xlsx"

REQUIRED = [
    "流水號", "申請序號", "發表日期", "申請者", "申請者科別",
    "申請項目", "論文類別", "獎金", "IF", "學門排行",
    "論文名稱", "發表之雜誌/期刊", "發表地點", "主管簽章",
    "作者1", "作者2", "作者3", "作者4", "作者5", "作者6",
]

def as_text(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    return str(value).strip()

def id_of(row):
    return (row[0].strip(), row[1].strip())

def parse_original(html):
    match = re.search(r"const RAW=(\{.*?\});\s*\nconst PHOTO=", html, re.S)
    if not match:
        raise ValueError("Cannot locate current published RAW records")
    return match, json.loads(match.group(1))

def main():
    if not SOURCE.exists():
        raise FileNotFoundError("Upload the full Excel to data/論文申請資料.xlsx")
    html = INDEX.read_text(encoding="utf-8")
    match, original = parse_original(html)
    if original["h"] != REQUIRED:
        raise ValueError("Published source columns do not match required schema")

    wb = load_workbook(SOURCE, data_only=True, read_only=True)
    rows = None
    for ws in wb.worksheets:
        iterator = ws.iter_rows(values_only=True)
        header = next(iterator, None)
        if not header:
            continue
        names = [as_text(v) for v in header]
        if not all(name in names for name in REQUIRED):
            continue
        indices = [names.index(k) for k in REQUIRED]
        rows = []
        for values in iterator:
            row = [as_text(values[i] if i < len(values) else None) for i in indices]
            if not any(row):
                continue
            if not row[0] or not row[1] or not row[3] or not row[10]:
                raise ValueError(f"Incomplete mandatory data at Excel row {len(rows)+2}")
            rows.append(row)
        break
    if rows is None:
        raise ValueError("Excel has no worksheet containing the original 20 headers")

    old = original["r"]
    old_by_id = {id_of(r): r for r in old}
    if len(old_by_id) != len(old):
        raise ValueError("Existing published records have duplicated IDs")
    new_by_id = {}
    for r in rows:
        key = id_of(r)
        if key in new_by_id:
            raise ValueError(f"Duplicate application ID in Excel: {key}")
        new_by_id[key] = r
    missing = set(old_by_id) - set(new_by_id)
    changed = [key for key in old_by_id.keys() & new_by_id.keys() if old_by_id[key] != new_by_id[key]]
    if missing or changed:
        raise ValueError(
            f"Update rejected: {len(missing)} historical rows missing, "
            f"{len(changed)} historical rows modified. "
            "Do not overwrite history. Review corrections separately."
        )
    if len(rows) < len(old):
        raise ValueError("Data shrink rejected")

    # Count award vs publication-fee applications separately.
    papers = sum("刊登費" not in r[5] for r in rows)
    fees = len(rows) - papers
    new_count = len(rows) - len(old)
    payload = json.dumps({"h": REQUIRED, "r": rows}, ensure_ascii=False, separators=(",", ":"))
    updated = html[:match.start(1)] + payload + html[match.end(1):]
    INDEX.write_text(updated, encoding="utf-8")
    print(f"Validated {len(rows)} rows: {papers} awards, {fees} publication fees; {new_count} new records.")
    print("No existing application was removed or modified. index.html updated.")

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
