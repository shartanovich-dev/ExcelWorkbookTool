"""Check that a code change did not change conversion results.

Compares every converted workbook in Baseline_output/ with the same-named file in Output/.
An .xlsx is a zip of XML parts and the zip stores write times, so we compare the content
of each part inside the files rather than the files' raw bytes.

Usage (from the project folder):
    python tools/compare_outputs.py                     # workbooks + diagnostic YAML
    python tools/compare_outputs.py run_console.txt     # also compare a saved console log
"""
import os
import sys
import zipfile

BASELINE_DIR = "Baseline_output"
OUTPUT_DIR = "Output"
YAML_NAME = "Cache_Diagnostic_Log.yaml"
BASELINE_CONSOLE = os.path.join(BASELINE_DIR, "baseline_console.txt")


def list_workbooks(folder):
    return sorted(f for f in os.listdir(folder)
                  if f.lower().endswith(".xlsx") and not f.startswith("~$"))


def compare_workbooks(path_a, path_b):
    """Return the names of parts that differ (empty list = identical)."""
    with zipfile.ZipFile(path_a) as za, zipfile.ZipFile(path_b) as zb:
        parts_a, parts_b = set(za.namelist()), set(zb.namelist())
        diffs = [f"missing: {p}" for p in sorted(parts_a - parts_b)]
        diffs += [f"extra: {p}" for p in sorted(parts_b - parts_a)]
        diffs += [p for p in sorted(parts_a & parts_b) if za.read(p) != zb.read(p)]
    return diffs


def same_file(path_a, path_b):
    with open(path_a, "rb") as fa, open(path_b, "rb") as fb:
        return fa.read() == fb.read()


def main():
    failures = 0
    baseline_files = list_workbooks(BASELINE_DIR)
    output_files = set(list_workbooks(OUTPUT_DIR))

    for name in baseline_files:
        if name not in output_files:
            print(f"MISSING    {name}")
            failures += 1
            continue
        diffs = compare_workbooks(os.path.join(BASELINE_DIR, name), os.path.join(OUTPUT_DIR, name))
        if diffs:
            print(f"DIFFERENT  {name}  {diffs}")
            failures += 1
        else:
            print(f"IDENTICAL  {name}")

    for name in sorted(output_files - set(baseline_files)):
        print(f"EXTRA      {name}  (in Output/ but not in baseline)")
        failures += 1

    checks = [(os.path.join(BASELINE_DIR, YAML_NAME), YAML_NAME, YAML_NAME)]
    if len(sys.argv) > 1:
        checks.append((BASELINE_CONSOLE, sys.argv[1], "console output"))
    for baseline_path, current_path, label in checks:
        if not (os.path.exists(baseline_path) and os.path.exists(current_path)):
            print(f"SKIPPED    {label}  (file not found)")
        elif same_file(baseline_path, current_path):
            print(f"IDENTICAL  {label}")
        else:
            print(f"DIFFERENT  {label}")
            failures += 1

    print()
    if failures:
        print(f"RESULT: {failures} difference(s) found")
        return 1
    print("RESULT: all identical")
    return 0


if __name__ == "__main__":
    sys.exit(main())
