import os
import glob
from app.dbrw_processor import process_workbook_xml
from app.mappings import load_ram_mappings

MAPPINGS_DIR = "./Mappings"
INPUT_DIR = "./Input"
OUTPUT_DIR = "./Output"

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("Loading Mapping files into RAM...")
    ram_mappings = load_ram_mappings(MAPPINGS_DIR)
    print(f"Loaded mappings for {len(ram_mappings)} cube(s): {list(ram_mappings.keys())}")

    input_files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
    for file_path in input_files:
        file_name = os.path.basename(file_path)

        if file_name.startswith("~$"):
            continue

        print(f"\nProcessing target file: {file_name}")

        with open(file_path, "rb") as f:
            input_bytes = f.read()

        processed_bytes = process_workbook_xml(input_bytes, ram_mappings)

        name_only, ext = os.path.splitext(file_name)
        new_file_name = f"{name_only}_Converted{ext}"
        output_path = os.path.join(OUTPUT_DIR, new_file_name)

        with open(output_path, "wb") as f:
            f.write(processed_bytes)

        print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
