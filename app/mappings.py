import glob
import os
import re

import openpyxl


# --- IRONCLAD NORMALIZER ---
def normalize_account_name(acct_str):
    if not acct_str:
        return ""
    s = str(acct_str).upper()
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s

def load_ram_mappings(mappings_dir):
    """Reads all Mapping_{cubename}.xlsx files from mappings_dir into memory."""
    ram_mappings = {}
    mapping_files = glob.glob(os.path.join(mappings_dir, "Mapping_*.xlsx"))

    for filepath in mapping_files:
        filename = os.path.basename(filepath)
        cube_name = filename.replace("Mapping_", "").replace(".xlsx", "").strip()

        wb = openpyxl.load_workbook(filepath, data_only=True)

        # --- TAB 1: Account_Mapping ---
        ws1 = wb["Account_Mapping"]
        headers = [cell.value for cell in ws1[1] if cell.value is not None]
        dz_headers = headers[2:] if len(headers) > 2 else []

        tab1_accounts = {}
        tab_exceptions = {}

        raw_mappings = {}
        for row in ws1.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue

            legacy_acct = normalize_account_name(row[0])
            wd_acct = str(row[1]).strip() if row[1] is not None and str(row[1]).strip() != "" else None
            dz_vals = [str(val).strip() if val is not None and str(val).strip() != "" else None for val in row[2:2 + len(dz_headers)]]

            while len(dz_vals) < len(dz_headers):
                dz_vals.append(None)

            if legacy_acct not in raw_mappings:
                raw_mappings[legacy_acct] = []

            raw_mappings[legacy_acct].append((wd_acct, dz_vals))

        # Route to 1-to-1 or Exception dictionaries
        for legacy_acct, mappings in raw_mappings.items():
            if len(mappings) == 1:
                tab1_accounts[legacy_acct] = mappings[0]
            else:
                ex_list = []
                for wd, dz in mappings:
                    cat_str = dz[0] if dz[0] is not None else ""
                    ex_list.append((wd, cat_str))
                tab_exceptions[legacy_acct] = ex_list

        # --- TAB 2: Context_Mapping (Unified Context & Element Mappings) ---
        ws2 = wb["Context_Mapping"]
        tab2_context = {}
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if row and row[0] is not None and row[1] is not None:
                tab2_context[str(row[0]).strip()] = str(row[1]).strip()

        # --- TAB 3: Blueprint_Mapping ---
        ws3 = wb["Blueprint_Mapping"]
        tab3_blueprint = str(ws3["B1"].value or "").strip()
        account_index = int(ws3["B2"].value or 1)

        ram_mappings[cube_name] = {
            'tab1_accounts': tab1_accounts,
            'tab_exceptions': tab_exceptions,
            'tab2_context': tab2_context,
            'tab3_blueprint': tab3_blueprint,
            'account_index': account_index,
            'dropzone_headers': dz_headers
        }

    return ram_mappings
