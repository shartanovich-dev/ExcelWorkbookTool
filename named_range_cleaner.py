"""Remove broken and unused named ranges (defined names) from Excel workbooks.

Removed:
  - Broken: the name's value contains #REF!
  - Unused: nothing in the workbook refers to the name (cell formulas, data validation,
    conditional formatting, charts, pivot tables, hyperlinks, or other kept names)
Always kept: names starting with any prefix in KEEP_PREFIXES (Excel built-ins, PAFE/TM1, ...)

Only the removed <definedName> entries change in xl/workbook.xml; every other part of the
workbook is copied unchanged. Originals are never modified.

Usage (from the project folder):  python named_range_cleaner.py
Reads .xlsx files from 'Named ranges/' and writes to 'Named ranges/Cleaned/':
    <name>_Cleaned.xlsx         the workbook without the removed names
    <name>_Removed_Names.csv    one row per removed name, with the reason
"""
import csv
import html
import os
import re
import zipfile
from io import BytesIO

INPUT_DIR = "Named ranges"
OUTPUT_DIR = os.path.join(INPUT_DIR, "Cleaned")

# Names starting with these (case-insensitive) are never removed
KEEP_PREFIXES = ["_xlnm.", "TM1"]
KEEP_EXACT = ["ID"]

DEFINED_NAMES_BLOCK = re.compile(r"<(?:\w+:)?definedNames\b[^>]*>(.*?)</(?:\w+:)?definedNames>", re.S)
DEFINED_NAME = re.compile(r"<(?:\w+:)?definedName\b([^>]*?)(?:/>|>(.*?)</(?:\w+:)?definedName>)", re.S)
SHEET_TAG = re.compile(r"<(?:\w+:)?sheet\b[^>]*\bname=\"([^\"]*)\"")
# Places inside sheets/charts/pivots where formulas or name references are stored
FORMULA_TEXT = re.compile(r"<(?:\w+:)?(?:f|formula|formula1|formula2)\b[^>]*>(.*?)</(?:\w+:)?(?:f|formula|formula1|formula2)>", re.S)
NAME_ATTRS = re.compile(r"\b(?:location|name|ref)=\"([^\"]*)\"")
WORD = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]*")


def get_attr(attrs, attr_name):
    m = re.search(r'\b' + attr_name + r'="([^"]*)"', attrs)
    return html.unescape(m.group(1)) if m else ""


def is_protected(name):
    upper = name.upper()
    return upper in (k.upper() for k in KEEP_EXACT) or any(upper.startswith(p.upper()) for p in KEEP_PREFIXES)


def collect_words(zin):
    """Every word that appears in a formula or name reference anywhere outside workbook.xml (lowercase)."""
    words = set()
    for part in zin.namelist():
        if not part.endswith(".xml"):
            continue
        if not part.startswith(("xl/worksheets/", "xl/charts/", "xl/pivotCache", "xl/pivotTables/")):
            continue
        text = zin.read(part).decode("utf-8", "ignore")
        for m in FORMULA_TEXT.finditer(text):
            words.update(w.lower() for w in WORD.findall(html.unescape(m.group(1))))
        if part.startswith(("xl/pivotCache", "xl/worksheets/")):     # pivot sources by name, hyperlinks to names
            for m in NAME_ATTRS.finditer(text):
                words.update(w.lower() for w in WORD.findall(html.unescape(m.group(1))))
    return words


def clean_workbook_xml(workbook_xml, used_words):
    """Return (new workbook.xml text, list of removed names, total names found)."""
    sheet_names = [html.unescape(n) for n in SHEET_TAG.findall(workbook_xml)]
    block = DEFINED_NAMES_BLOCK.search(workbook_xml)
    if not block:
        return workbook_xml, [], 0

    entries = []
    for m in DEFINED_NAME.finditer(block.group(1)):
        attrs = m.group(1)
        entries.append({
            "xml": m.group(0),
            "name": get_attr(attrs, "name"),
            "local_id": get_attr(attrs, "localSheetId"),
            "hidden": get_attr(attrs, "hidden") in ("1", "true"),
            "refers_to": html.unescape(m.group(2) or ""),
        })

    # 1. Broken names
    for e in entries:
        e["reason"] = "Broken (#REF!)" if "#REF!" in e["refers_to"].upper() else None

    # 2. Kept = protected or used; names referenced by kept names also count as used (repeat until stable)
    used = set(used_words)
    while True:
        kept_now = [e for e in entries if e["reason"] is None and (is_protected(e["name"]) or e["name"].lower() in used)]
        new_words = {w.lower() for e in kept_now for w in WORD.findall(e["refers_to"])} - used
        if not new_words:
            break
        used |= new_words
    for e in entries:
        if e["reason"] is None and not is_protected(e["name"]) and e["name"].lower() not in used:
            e["reason"] = "Unused"

    removed, kept_xml = [], []
    for e in entries:
        if e["reason"] is None:
            kept_xml.append(e["xml"])
            continue
        lid = e["local_id"]
        scope = sheet_names[int(lid)] if lid.isdigit() and int(lid) < len(sheet_names) else "Workbook"
        removed.append({"Name": e["name"], "Reason": e["reason"], "Scope": scope,
                        "Hidden": "Yes" if e["hidden"] else "No", "Refers To": e["refers_to"]})

    if not removed:
        return workbook_xml, [], len(entries)
    # Rebuild only the <definedNames> block; drop it entirely if nothing is left
    new_block = block.group(0).replace(block.group(1), "".join(kept_xml)) if kept_xml else ""
    new_xml = workbook_xml[:block.start()] + new_block + workbook_xml[block.end():]
    return new_xml, removed, len(entries)


def clean_workbook(input_bytes):
    """Return (cleaned workbook bytes, list of removed names, total names found)."""
    output_buffer = BytesIO()
    with zipfile.ZipFile(BytesIO(input_bytes)) as zin, \
         zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
        used_words = collect_words(zin)
        new_xml, removed, total = clean_workbook_xml(zin.read("xl/workbook.xml").decode("utf-8"), used_words)
        for item in zin.infolist():                     # same parts, same order, same compression
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = new_xml.encode("utf-8")
            zout.writestr(item, data)
    return output_buffer.getvalue(), removed, total


def write_report(path, removed):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:   # utf-8-sig opens cleanly in Excel
        writer = csv.DictWriter(f, fieldnames=["Name", "Reason", "Scope", "Hidden", "Refers To"])
        writer.writeheader()
        writer.writerows(removed)


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    files = sorted(f for f in os.listdir(INPUT_DIR)
                   if f.lower().endswith(".xlsx") and not f.startswith("~$")
                   and os.path.isfile(os.path.join(INPUT_DIR, f)))
    for file_name in files:
        with open(os.path.join(INPUT_DIR, file_name), "rb") as f:
            input_bytes = f.read()
        cleaned_bytes, removed, total = clean_workbook(input_bytes)

        base = os.path.splitext(file_name)[0]
        with open(os.path.join(OUTPUT_DIR, f"{base}_Cleaned.xlsx"), "wb") as f:
            f.write(cleaned_bytes)
        write_report(os.path.join(OUTPUT_DIR, f"{base}_Removed_Names.csv"), removed)

        broken = sum(1 for r in removed if r["Reason"].startswith("Broken"))
        print(f"{file_name}: {total:,} names -> {total - len(removed):,} kept "
              f"(removed {broken:,} broken, {len(removed) - broken:,} unused)")


if __name__ == "__main__":
    main()
