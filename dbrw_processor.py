import re
import xml.etree.ElementTree as ET
import zipfile
import copy
from io import BytesIO

# XML Namespaces
NS_URL = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS = {'s': NS_URL}
ET.register_namespace('', NS_URL)

def normalize_account_name(acct_str):
    if not acct_str:
        return ""
    s = str(acct_str).upper()
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s

def get_col_letter(col_idx):
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def col_letter_to_idx(col_let):
    idx = 0
    for char in col_let.upper():
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx

def parse_col_letter(cell_ref):
    match = re.match(r"([A-Z]+)", cell_ref, re.IGNORECASE)
    return match.group(1).upper() if match else "A"

def parse_row_num(cell_ref):
    match = re.match(r"[A-Z]+(\d+)", cell_ref, re.IGNORECASE)
    return match.group(1) if match else "1"

def split_dbrw_params(param_str):
    params = []
    current = []
    in_quotes = False
    quote_char = ''
    
    for char in param_str:
        if char in ('"', "'"):
            if not in_quotes:
                in_quotes = True
                quote_char = char
            elif char == quote_char:
                in_quotes = False
            current.append(char)
        elif char == ',' and not in_quotes:
            params.append("".join(current).strip())
            current = []
        else:
            current.append(char)
    if current:
        params.append("".join(current).strip())
    return params

def load_shared_strings(input_zip):
    shared_strings = []
    try:
        sst_xml = input_zip.read('xl/sharedStrings.xml')
        sst_tree = ET.fromstring(sst_xml)
        for si in sst_tree.findall('.//s:si', NS):
            text_parts = []
            for t in si.findall('.//s:t', NS):
                if t.text:
                    text_parts.append(t.text)
            shared_strings.append("".join(text_parts))
    except KeyError:
        pass 
    return shared_strings

def get_cell_value(sheet_tree, cell_ref, shared_strings):
    clean_ref = cell_ref.replace('$', '').strip()
    for cell in sheet_tree.findall('.//s:c', NS):
        if cell.attrib.get('r') == clean_ref:
            cell_type = cell.attrib.get('t')
            
            if cell_type == 'inlineStr':
                is_tag = cell.find('s:is/s:t', NS)
                return is_tag.text.strip() if is_tag is not None and is_tag.text else None
                
            v_tag = cell.find('s:v', NS)
            if v_tag is not None and v_tag.text:
                if cell_type == 's':
                    idx = int(v_tag.text)
                    return shared_strings[idx].strip() if idx < len(shared_strings) else v_tag.text.strip()
                return v_tag.text.strip()
    return None

def process_workbook_xml(input_file_bytes, ram_mappings):
    input_zip = zipfile.ZipFile(BytesIO(input_file_bytes), 'r')
    output_buffer = BytesIO()
    output_zip = zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED)

    structural_files = {
        '[Content_Types].xml': None,
        'xl/workbook.xml': None,
        'xl/_rels/workbook.xml.rels': None,
        'docProps/app.xml': None
    }
    
    shared_strings = load_shared_strings(input_zip)
    workbook_xml_raw = input_zip.read('xl/workbook.xml')
    wb_tree = ET.fromstring(workbook_xml_raw)
    rels_xml_raw = input_zip.read('xl/_rels/workbook.xml.rels')
    rels_tree = ET.fromstring(rels_xml_raw)
    
    rel_map = {
        elem.attrib['Id']: elem.attrib['Target'] 
        for elem in rels_tree.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
    }

    sheet_info = []
    for sheet in wb_tree.findall('.//s:sheet', NS):
        r_id = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        tab_name = sheet.attrib.get('name', 'Unknown')
        if r_id in rel_map:
            target = rel_map[r_id]
            if not target.startswith('xl/'):
                target = 'xl/' + target
            sheet_info.append((target, tab_name))

    sheet_paths = [info[0] for info in sheet_info]

    for item in input_zip.infolist():
        if item.filename == 'xl/calcChain.xml':
            continue
        if item.filename in structural_files:
            structural_files[item.filename] = input_zip.read(item.filename)
            continue
        if item.filename not in sheet_paths:
            output_zip.writestr(item.filename, input_zip.read(item.filename))

    audit_log = {}

    for sheet_path, tab_name in sheet_info:
        # Extract sheet number to cross-reference with Excel repair logs
        sheet_num_match = re.search(r'sheet(\d+)\.xml', sheet_path, re.IGNORECASE)
        sheet_num = sheet_num_match.group(1) if sheet_num_match else "Unknown"
        
        print(f"\n[Sheet {sheet_num}]: '{tab_name}' ({sheet_path})")
        raw_sheet_bytes = input_zip.read(sheet_path)
        raw_xml_str = raw_sheet_bytes.decode('utf-8')

        sheet_data_match = re.search(r'(<sheetData[^>]*>)(.*?)(</sheetData>)', raw_xml_str, flags=re.DOTALL)
        
        if not sheet_data_match:
            output_zip.writestr(sheet_path, raw_sheet_bytes)
            continue

        header_text = raw_xml_str[:sheet_data_match.start(2)]
        inner_text = sheet_data_match.group(2)
        footer_text = raw_xml_str[sheet_data_match.end(2):]

        ns_declarations = " ".join(re.findall(r'xmlns(?::[a-zA-Z0-9\-]+)?="[^"]*"', header_text))
        dummy_xml = f'<root {ns_declarations}>{inner_text}</root>'
        
        try:
            sheet_tree = ET.fromstring(dummy_xml)
        except ET.ParseError:
            output_zip.writestr(sheet_path, raw_sheet_bytes)
            continue

        dbrw_formula_text = None
        start_row = 1
        max_row = 0
        
        for cell in sheet_tree.findall('.//s:c', NS):
            ref = cell.attrib.get('r', '')
            row_num = int(parse_row_num(ref)) if ref else 999
            if row_num > max_row and row_num != 999:
                max_row = row_num
            if row_num > 30:
                continue
                
            f_tag = cell.find('s:f', NS)
            if f_tag is not None and f_tag.text and 'DBRW' in f_tag.text.upper() and not dbrw_formula_text:
                dbrw_formula_text = f_tag.text
                start_row = row_num

        if not dbrw_formula_text:
            output_zip.writestr(sheet_path, raw_sheet_bytes)
            continue

        audit_log[tab_name] = {'Evaluated': 0, 'Converted': 0, 'Exceptions': 0, 'Incomplete': 0, 'Unmapped': 0, 'Duplicates': 0}
        seen_wd_combos = set()

        params_match = re.search(r'DBRW\s*\((.*)\)', dbrw_formula_text, re.IGNORECASE)
        extracted_cube = None
        params = []
        
        if params_match:
            params = split_dbrw_params(params_match.group(1))
            if params:
                param_1 = params[0].strip()
                if re.match(r'^\$?[A-Z]+\$?\d+$', param_1, re.IGNORECASE):
                    extracted_cube = get_cell_value(sheet_tree, param_1, shared_strings)
                else:
                    extracted_cube = param_1.strip('"').strip("'")

        matched_cube_key = None
        if extracted_cube:
            if extracted_cube in ram_mappings:
                matched_cube_key = extracted_cube
            else:
                for ram_key in ram_mappings.keys():
                    if ram_key.lower() in extracted_cube.lower():
                        matched_cube_key = ram_key
                        break

        extracted_cube = matched_cube_key

        if not extracted_cube or extracted_cube not in ram_mappings:
            output_zip.writestr(sheet_path, raw_sheet_bytes)
            continue

        mapping_data = ram_mappings[extracted_cube]
        tab1_accounts = mapping_data['tab1_accounts']
        tab_exceptions = mapping_data.get('tab_exceptions', {})
        tab2_context  = mapping_data['tab2_context']
        tab3_blueprint = mapping_data['tab3_blueprint']
        account_idx    = mapping_data['account_index']
        dz_headers     = mapping_data['dropzone_headers']

        # Determine MAX COLUMN to calculate offsets for Tag and Dropzones
        max_col_idx = 0
        for cell in sheet_tree.findall('.//s:c', NS):
            col_let = parse_col_letter(cell.attrib.get('r', 'A'))
            idx = col_letter_to_idx(col_let)
            if idx > max_col_idx:
                max_col_idx = idx

        # Column offsets: VBA tag is exactly at +1, Dropzones start at +2
        vba_col_idx = max_col_idx + 1
        start_dz_idx = max_col_idx + 2

        dz_col_map = {}
        for i, header_name in enumerate(dz_headers):
            dz_col_map[header_name] = get_col_letter(start_dz_idx + i)

        account_col_letter = None
        for cell in sheet_tree.findall('.//s:c', NS):
            f_tag = cell.find('s:f', NS)
            if f_tag is not None and f_tag.text and 'DBRW' in f_tag.text.upper():
                p_match = re.search(r'DBRW\s*\((.*)\)', f_tag.text, re.IGNORECASE)
                if p_match:
                    p_list = split_dbrw_params(p_match.group(1))
                    if len(p_list) >= account_idx:
                        acct_ref = p_list[account_idx - 1].replace('$', '')
                        account_col_letter = parse_col_letter(acct_ref)
                        break

        target_header_refs = []
        for param in params:
            if re.match(r'^\$?[A-Z]+\$?\d+$', param, re.IGNORECASE):
                target_header_refs.append(param.replace('$', '').strip().upper())
                
        for cell in sheet_tree.findall('.//s:c', NS):
            cell_ref = cell.attrib.get('r', '').upper()
            if cell_ref in target_header_refs:
                f_tag = cell.find('s:f', NS)
                v_tag = cell.find('s:v', NS)
                cell_type = cell.attrib.get('t')
                
                original_text = None
                is_formula = False
                
                if f_tag is not None and f_tag.text:
                    original_text = f_tag.text
                    is_formula = True
                else:
                    if cell_type == 'inlineStr':
                        texts = cell.findall('.//s:t', NS)
                        if texts:
                            original_text = "".join(t.text for t in texts if t.text)
                    elif cell_type == 's' and v_tag is not None and v_tag.text:
                        idx = int(v_tag.text)
                        if idx < len(shared_strings):
                            original_text = shared_strings[idx]
                    elif v_tag is not None and cell_type != 'n':
                        original_text = v_tag.text
                
                if not original_text or not isinstance(original_text, str):
                    continue
                    
                new_text = original_text
                for find_str, replace_str in tab2_context.items():
                    if find_str in new_text:
                        new_text = new_text.replace(find_str, replace_str)
                        
                if new_text != original_text:
                    if is_formula:
                        f_tag.text = new_text
                        if v_tag is not None:
                            cell.remove(v_tag)
                    else:
                        cell.attrib['t'] = 'inlineStr'
                        for child in list(cell):
                            cell.remove(child)
                        
                        new_is = ET.SubElement(cell, f'{{{NS_URL}}}is')
                        new_t = ET.SubElement(new_is, f'{{{NS_URL}}}t')
                        new_t.text = new_text

        cloned_rows_buffer = []

        for row in sheet_tree.findall('s:row', NS):
            row_num_str = row.attrib.get('r')
            if row_num_str and int(row_num_str) < start_row:
                continue
                
            row_num = row_num_str
            account_val = None
            acct_cell_elem = None

            if account_col_letter:
                target_ref = f"{account_col_letter}{row_num}"
                for cell in row.findall('s:c', NS):
                    if cell.attrib.get('r') == target_ref:
                        acct_cell_elem = cell
                        v_tag = cell.find('s:v', NS)
                        cell_type = cell.attrib.get('t')
                        
                        if cell_type == 'inlineStr':
                            texts = cell.findall('.//s:t', NS)
                            if texts:
                                account_val = "".join(t.text for t in texts if t.text)
                        elif cell_type == 's' and v_tag is not None and v_tag.text:
                            idx = int(v_tag.text)
                            if idx < len(shared_strings):
                                account_val = shared_strings[idx]
                            else:
                                account_val = v_tag.text
                        elif v_tag is not None:
                            account_val = v_tag.text
                        break

            base_row_snapshot = None
            normalized_account_val = normalize_account_name(account_val) if account_val else None
            is_exception = (normalized_account_val in tab_exceptions) if normalized_account_val else False
            
            if is_exception:
                base_row_snapshot = copy.deepcopy(row)

            for cell in row.findall('s:c', NS):
                f_tag = cell.find('s:f', NS)
                if f_tag is not None and f_tag.text and 'DBRW' in f_tag.text.upper():
                    old_formula = f_tag.text
                    p_match = re.search(r'DBRW\s*\((.*)\)', old_formula, re.IGNORECASE)
                    
                    if p_match:
                        extracted_params = split_dbrw_params(p_match.group(1))
                        new_blueprint = tab3_blueprint
                        for idx, p_val in enumerate(extracted_params):
                            new_blueprint = new_blueprint.replace(f"{{Param_{idx}}}", p_val)
                            
                        for header_name, col_let in dz_col_map.items():
                            new_blueprint = new_blueprint.replace(f"{{{header_name}}}", f"${col_let}{row_num}")
                        
                        for old_val, new_val in tab2_context.items():
                            new_blueprint = new_blueprint.replace(old_val, new_val)

                        updated_formula = re.sub(
                            r'_xll\.DBRW\s*\([^)]*\)|DBRW\s*\([^)]*\)', 
                            new_blueprint, 
                            old_formula, 
                            flags=re.IGNORECASE
                        )

                        updated_formula = updated_formula.lstrip('=')
                        f_tag.text = updated_formula
                        f_tag.attrib.pop('t', None)
                        f_tag.attrib.pop('si', None)
                        f_tag.attrib.pop('ref', None)

                        v_tag = cell.find('s:v', NS)
                        if v_tag is not None:
                            cell.remove(v_tag)

            if account_val:
                audit_log[tab_name]['Evaluated'] += 1
                is_scenario_1 = False
                
                if is_exception:
                    audit_log[tab_name]['Exceptions'] += 1
                    is_scenario_1 = True
                    exception_list = tab_exceptions[normalized_account_val]
                    
                    wd_account_1, cat_1 = exception_list[0]
                    wd_account = wd_account_1
                    dz_values = [cat_1] + [""] * (len(dz_headers) - 1)
                    
                    # --- NEW DUPLICATE CHECK FOR EXCEPTION BASE ROW ---
                    combo = (wd_account, tuple(dz_values))
                    if combo in seen_wd_combos:
                        dz_values = ["[DUPLICATE]"] * len(dz_headers)
                        is_scenario_1 = False 
                        audit_log[tab_name]['Duplicates'] += 1
                    else:
                        seen_wd_combos.add(combo)
                    # --------------------------------------------------
                    
                    for split_wd, split_cat in exception_list[1:]:
                        max_row += 1
                        new_row_num = str(max_row)
                        cloned_row = copy.deepcopy(base_row_snapshot)
                        cloned_row.attrib['r'] = new_row_num
                        
                        # 1. Update Clone Coordinates
                        for cell in cloned_row.findall('s:c', NS):
                            old_ref = cell.attrib.get('r', '')
                            if old_ref:
                                col_let = parse_col_letter(old_ref)
                                cell.attrib['r'] = f"{col_let}{new_row_num}"
                                
                        # 2. Inject Clone Account Name
                        if account_col_letter:
                            target_ref = f"{account_col_letter}{new_row_num}"
                            acct_cell = None
                            for c in cloned_row.findall('s:c', NS):
                                if c.attrib.get('r') == target_ref:
                                    acct_cell = c
                                    break
                            if not acct_cell:
                                acct_cell = ET.Element(f'{{{NS_URL}}}c', {'r': target_ref})
                                cloned_row.append(acct_cell)
                                
                            acct_cell.attrib['t'] = 'inlineStr'
                            for child in list(acct_cell): acct_cell.remove(child)
                            ET.SubElement(ET.SubElement(acct_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str(split_wd)
                            
                        # 3. Inject Clone Dropzones
                        dz_values_clone = [split_cat] + [""] * (len(dz_headers) - 1)
                        
                        # --- NEW DUPLICATE CHECK FOR CLONED ROW ---
                        combo_clone = (split_wd, tuple(dz_values_clone))
                        if combo_clone in seen_wd_combos:
                            dz_values_clone = ["[DUPLICATE]"] * len(dz_headers)
                            audit_log[tab_name]['Duplicates'] += 1
                        else:
                            seen_wd_combos.add(combo_clone)
                        # ------------------------------------------
                        
                        for h_idx, dz_val in enumerate(dz_values_clone):
                            target_dz_col = dz_col_map[dz_headers[h_idx]]
                            dz_cell_ref = f"{target_dz_col}{new_row_num}"
                            for c in cloned_row.findall('s:c', NS):
                                if c.attrib.get('r') == dz_cell_ref:
                                    cloned_row.remove(c)
                            dz_cell = ET.Element(f'{{{NS_URL}}}c', {'r': dz_cell_ref, 't': 'inlineStr'})
                            ET.SubElement(ET.SubElement(dz_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str(dz_val)
                            cloned_row.append(dz_cell)
                            
                        # 4. Update Clone Formulas & Wipe Shared Formulas
                        for cell in cloned_row.findall('s:c', NS):
                            f_tag = cell.find('s:f', NS)
                            if f_tag is not None:
                                # Wipe out Excel shared formulas (e.g. SUMs) entirely in the clone to avoid corruption
                                if not (f_tag.text and 'DBRW' in f_tag.text.upper()):
                                    cell.remove(f_tag)
                                    v_tag = cell.find('s:v', NS)
                                    if v_tag is not None:
                                        cell.remove(v_tag)
                                else:
                                    # Handle DBRW Updates normally
                                    old_formula = f_tag.text
                                    p_match = re.search(r'DBRW\s*\((.*)\)', old_formula, re.IGNORECASE)
                                    if p_match:
                                        extracted_params = split_dbrw_params(p_match.group(1))
                                        new_blueprint = tab3_blueprint
                                        for idx, p_val in enumerate(extracted_params):
                                            p_val_renumbered = re.sub(r'(\$?[A-Z]+\$?)' + str(row_num) + r'\b', r'\g<1>' + new_row_num, p_val, flags=re.IGNORECASE)
                                            new_blueprint = new_blueprint.replace(f"{{Param_{idx}}}", p_val_renumbered)
                                            
                                        for header_name, col_let in dz_col_map.items():
                                            new_blueprint = new_blueprint.replace(f"{{{header_name}}}", f"${col_let}{new_row_num}")
                                        
                                        for old_val, new_val in tab2_context.items():
                                            new_blueprint = new_blueprint.replace(old_val, new_val)

                                        updated_formula = re.sub(r'_xll\.DBRW\s*\([^)]*\)|DBRW\s*\([^)]*\)', new_blueprint, old_formula, flags=re.IGNORECASE)
                                        updated_formula = updated_formula.lstrip('=')
                                        f_tag.text = updated_formula
                                        f_tag.attrib.pop('t', None)
                                        f_tag.attrib.pop('si', None)
                                        f_tag.attrib.pop('ref', None)

                                        v_tag = cell.find('s:v', NS)
                                        if v_tag is not None:
                                            cell.remove(v_tag)

                        # 5. Append VBA Tag dynamically at Max Column + 1
                        vba_col_let = get_col_letter(vba_col_idx)
                        vba_cell_ref = f"{vba_col_let}{new_row_num}"
                        vba_cell = ET.Element(f'{{{NS_URL}}}c', {'r': vba_cell_ref, 't': 'inlineStr'})
                        ET.SubElement(ET.SubElement(vba_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = f"[INSERT_BELOW_{row_num}]"
                        cloned_row.append(vba_cell)
                        
                        # 6. Sort cells & Add to buffer
                        cells_c = [c for c in cloned_row if c.tag == f'{{{NS_URL}}}c']
                        cells_c.sort(key=lambda c: col_letter_to_idx(parse_col_letter(c.attrib.get('r', 'A'))))
                        for c in cells_c: cloned_row.remove(c)
                        cloned_row.extend(cells_c)
                        cloned_rows_buffer.append(cloned_row)

                elif normalized_account_val in tab1_accounts:
                    wd_account, dz_values_raw = tab1_accounts[normalized_account_val]
                    
                    if wd_account is None or str(wd_account).strip() == "":
                        dz_values = ["unused"] * len(dz_headers)
                        audit_log[tab_name]['Incomplete'] += 1
                        
                    elif str(wd_account).strip().lower() == "exception" or any(str(v).strip().lower() == "exception" for v in dz_values_raw if v is not None):
                        dz_values = [str(v) if v is not None else "" for v in dz_values_raw]
                        audit_log[tab_name]['Exceptions'] += 1
                        
                    else:
                        is_scenario_1 = True
                        dz_values = [str(v) if v is not None else "" for v in dz_values_raw]
                        
                        combo = (wd_account, tuple(dz_values))
                        if combo in seen_wd_combos:
                            dz_values = ["[DUPLICATE]"] * len(dz_headers)
                            is_scenario_1 = False 
                            audit_log[tab_name]['Duplicates'] += 1
                        else:
                            seen_wd_combos.add(combo)
                            audit_log[tab_name]['Converted'] += 1
                else:
                    dz_values = ["Unmapped"] * len(dz_headers)
                    audit_log[tab_name]['Unmapped'] += 1

                if is_scenario_1 and acct_cell_elem is not None:
                    acct_cell_elem.attrib['t'] = 'inlineStr'
                    for child in list(acct_cell_elem):
                        acct_cell_elem.remove(child)
                    
                    is_elem = ET.SubElement(acct_cell_elem, f'{{{NS_URL}}}is')
                    t_elem = ET.SubElement(is_elem, f'{{{NS_URL}}}t')
                    t_elem.text = str(wd_account)

                for h_idx, dz_val in enumerate(dz_values):
                    target_dz_col = dz_col_map[dz_headers[h_idx]]
                    dz_cell_ref = f"{target_dz_col}{row_num}"
                    
                    dz_cell = ET.Element(f'{{{NS_URL}}}c', {
                        'r': dz_cell_ref,
                        't': 'inlineStr'
                    })
                    dz_is = ET.SubElement(dz_cell, f'{{{NS_URL}}}is')
                    dz_t = ET.SubElement(dz_is, f'{{{NS_URL}}}t')
                    dz_t.text = str(dz_val)
                    row.append(dz_cell)

            cells = [c for c in row if c.tag == f'{{{NS_URL}}}c']
            cells.sort(key=lambda c: col_letter_to_idx(parse_col_letter(c.attrib.get('r', 'A'))))
            for c in cells:
                row.remove(c)
            row.extend(cells)

        for clone in cloned_rows_buffer:
            sheet_tree.append(clone)

        # --- UPDATED CONSOLE LOGGING OUTPUT ---
        if tab_name in audit_log:
            stats = audit_log[tab_name]
            print(f"  ├─ Evaluated:  {stats['Evaluated']}")
            print(f"  ├─ Converted:  {stats['Converted']}")
            print(f"  ├─ Exceptions: {stats['Exceptions']} (Triggered {len(cloned_rows_buffer)} clones)")
            print(f"  ├─ Duplicates: {stats['Duplicates']}")
            print(f"  └─ Unmapped:   {stats['Unmapped']}")

        modified_inner_str = ""
        for child in sheet_tree:
            modified_inner_str += ET.tostring(child, encoding='unicode')
        
        modified_inner_str = modified_inner_str.replace('ns0:', '').replace(':ns0', '')
        final_xml_str = header_text + modified_inner_str + footer_text
        output_zip.writestr(sheet_path, final_xml_str.encode('utf-8'))


    if structural_files['xl/workbook.xml'] and audit_log:
        try:
            def escape_xml(text):
                s = str(text)
                s = s.replace("&amp;", "&") 
                return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")

            wb_data = structural_files['xl/workbook.xml'].decode('utf-8')
            sheet_ids = [int(m) for m in re.findall(r'sheetId="(\d+)"', wb_data)]
            new_sheet_id = max(sheet_ids) + 1 if sheet_ids else 99
            
            rels_data = structural_files['xl/_rels/workbook.xml.rels'].decode('utf-8')
            r_ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_data)]
            max_r_id = max(r_ids) if r_ids else 99
            new_r_id = f"rId{max_r_id + 1}"

            new_sheet_node = f'<sheet name="Conversion Audit" sheetId="{new_sheet_id}" r:id="{new_r_id}"/>'
            wb_data = wb_data.replace('</sheets>', f'{new_sheet_node}</sheets>')
            
            new_rel_node = f'<Relationship Id="{new_r_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet_audit_{new_sheet_id}.xml"/>'
            rels_data = rels_data.replace('</Relationships>', f'{new_rel_node}</Relationships>')

            ct_data = structural_files['[Content_Types].xml'].decode('utf-8')
            new_override = f'<Override PartName="/xl/worksheets/sheet_audit_{new_sheet_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            ct_data = ct_data.replace('</Types>', f'{new_override}</Types>')

            output_zip.writestr('xl/workbook.xml', wb_data.encode('utf-8'))
            output_zip.writestr('xl/_rels/workbook.xml.rels', rels_data.encode('utf-8'))
            output_zip.writestr('[Content_Types].xml', ct_data.encode('utf-8'))
            if structural_files.get('docProps/app.xml'):
                output_zip.writestr('docProps/app.xml', structural_files['docProps/app.xml'])

            audit_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            audit_xml += f'<worksheet xmlns="{NS_URL}">\n'
            audit_xml += '<sheetData>\n'
            
            headers = ['Sheet Name', 'Total Evaluated', 'Converted (Success)', 'Exceptions', 'Incomplete Match', 'Unmapped (Orphans)', 'Duplicates Suppressed']
            audit_xml += '<row r="1">\n'
            for col_idx, header_val in enumerate(headers, start=1):
                col_let = get_col_letter(col_idx)
                safe_header = escape_xml(header_val)
                audit_xml += f'<c r="{col_let}1" t="inlineStr"><is><t>{safe_header}</t></is></c>\n'
            audit_xml += '</row>\n'

            row_num = 2
            for tab_n, stats in audit_log.items():
                audit_xml += f'<row r="{row_num}">\n'
                
                col_let = get_col_letter(1)
                safe_tab_name = escape_xml(tab_n)
                audit_xml += f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{safe_tab_name}</t></is></c>\n'
                
                numeric_stats = [
                    stats['Evaluated'],
                    stats['Converted'],
                    stats['Exceptions'],
                    stats['Incomplete'],
                    stats['Unmapped'],
                    stats['Duplicates']
                ]
                for i, num_val in enumerate(numeric_stats, start=2):
                    col_let = get_col_letter(i)
                    audit_xml += f'<c r="{col_let}{row_num}"><v>{num_val}</v></c>\n'

                audit_xml += '</row>\n'
                row_num += 1

            audit_xml += '</sheetData>\n</worksheet>'
            output_zip.writestr(f'xl/worksheets/sheet_audit_{new_sheet_id}.xml', audit_xml.encode('utf-8'))
            print(f"\n  ├─ Successfully Injected 'Conversion Audit' tab ({len(audit_log)} sheets logged).")

        except Exception as e:
            print(f"\n  └─ WARNING: Failed to inject Audit XML. Error: {e}")
            for f_name, f_data in structural_files.items():
                if f_data:
                    output_zip.writestr(f_name, f_data)
    else:
        for f_name, f_data in structural_files.items():
            if f_data:
                output_zip.writestr(f_name, f_data)

    output_zip.close()
    return output_buffer.getvalue()