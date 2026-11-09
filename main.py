import os
from app import local_store
from app.config import MAPPINGS_DIR
from app.dbrw_processor import process_workbook_xml
from app.mappings import load_ram_mappings

def main():
    local_store.ensure_output_folder()
    print("Loading Mapping files into RAM...")
    ram_mappings = load_ram_mappings(MAPPINGS_DIR)
    print(f"Loaded mappings for {len(ram_mappings)} cube(s): {list(ram_mappings.keys())}")

    for file_path in local_store.list_input_files():
        file_name = os.path.basename(file_path)

        print(f"\nProcessing target file: {file_name}")

        input_bytes = local_store.read_bytes(file_path)

        processed_bytes = process_workbook_xml(input_bytes, ram_mappings)

        name_only, ext = os.path.splitext(file_name)
        new_file_name = f"{name_only}_Converted{ext}"
        output_path = local_store.write_output(new_file_name, processed_bytes)

        print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
