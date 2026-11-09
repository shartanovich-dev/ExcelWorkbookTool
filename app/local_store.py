"""File handling for running locally (Input/, Output/ folders on this PC)."""
import glob
import os

from app.config import INPUT_DIR, OUTPUT_DIR


def ensure_output_folder():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

def list_input_files():
    """Paths of the .xlsx files in INPUT_DIR, skipping Excel's ~$ lock files."""
    input_files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
    return [p for p in input_files if not os.path.basename(p).startswith("~$")]

def read_bytes(file_path):
    with open(file_path, "rb") as f:
        return f.read()

def write_output(file_name, data):
    """Write data to OUTPUT_DIR/file_name and return the path written."""
    output_path = os.path.join(OUTPUT_DIR, file_name)
    with open(output_path, "wb") as f:
        f.write(data)
    return output_path
