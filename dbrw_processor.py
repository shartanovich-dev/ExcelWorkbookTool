import re
import xml.etree.ElementTree as ET
import zipfile
import copy
from io import BytesIO

# XML Namespaces
NS_URL = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS = {'s': NS_URL}
ET.register_namespace('', NS_URL)

# --- BULLETPROOF EXCEL REGEX PATTERNS ---
DBRW_EXTRACT_RGX = r'(?:_xll\.)?DBRW\s*\(((?:[^)("\']|"[^"]*"|\'[^\']*\'|\([^)]*\))*)\)'
DBRW_SUB_RGX = r'@?(?:_xll\.)?DBRW\s*\((?:[^)("\']|"[^"]*"|\'[^\']*\'|\([^)]*\))*\)'

# --- UTILITY HELPER FUNCTIONS ---
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

    ram_sheets = {}       
    sheet_wrappers = {}   
    sheet_path_map = {}
    raw_bytes_map = {}    
    audit_log = {}        

    print("  ├─ Loading all XML sheets into memory...")
    for sheet_path, tab_name in sheet_info:
        raw_sheet_bytes = input_zip.read(sheet_path)
        raw_bytes_map[tab_name] = raw_sheet_bytes
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
            ram_sheets[tab_name] = sheet_tree
            sheet_wrappers[tab_name] = (header_text, footer_text)
            sheet_path_map[tab_name] = sheet_path
            
            audit_log[tab_name] = {
                'Acct_Col': 'N/A', 'DZ_Cols': 'N/A',
                'Evaluated': 0, 'Converted': 0, 'Exceptions': 0, 'Incomplete': 0, 'Unmapped': 0, 'Duplicates': 0
            }
        except ET.ParseError:
            output_zip.writestr(sheet_path, raw_sheet_bytes)

    cube_cache = {}
    header_cache = {}
    account_cache = {}
    external_tab_cache = {}

    for tab_name, sheet_tree in ram_sheets.items():
        sheet_path = sheet_path_map[tab_name]
        s_num_match = re.search(r'sheet(\d+)\.xml', sheet_path, re.IGNORECASE)
        s_num = s_num_match.group(1) if s_num_match else "?"

        has_dbrw = False
        for f_node in sheet_tree.findall('.//s:f', NS):
            if f_node.text and 'DBRW' in f_node.text.upper():
                has_dbrw = True
                break
                
        if not has_dbrw:
            output_zip.writestr(sheet_path, raw_bytes_map[tab_name])
            continue
            
        print(f"\n  ├─ [Processing Tab]: '{tab_name}' (Sheet {s_num})")
        
        seen_wd_combos = {}
        popped_si_values = set()
        cloned_rows_buffer = []
        
        max_col_idx = 0
        for cell in sheet_tree.findall('.//s:c', NS):
            col_let = parse_col_letter(cell.attrib.get('r', 'A'))
            idx = col_letter_to_idx(col_let)
            if idx > max_col_idx:
                max_col_idx = idx
                
        vba_col_idx = max_col_idx + 1
        start_dz_idx = max_col_idx + 2

        max_row = 0
        for r_elem in sheet_tree.findall('.//s:row', NS):
            r_attr = r_elem.attrib.get('r')
            if r_attr and r_attr.isdigit():
                if int(r_attr) > max_row:
                    max_row = int(r_attr)

        for row in sheet_tree.findall('.//s:row', NS):
            row_num_str = row.attrib.get('r')
            if not row_num_str:
                continue
            row_num = row_num_str
            
            for cell in row.findall('s:c', NS):
                f_tag = cell.find('s:f', NS)
                
                if f_tag is not None and f_tag.text and 'DBRW' in f_tag.text.upper():
                    old_formula = f_tag.text
                    
                    p_match = re.search(DBRW_EXTRACT_RGX, old_formula, re.IGNORECASE)
                    if not p_match:
                        continue
                        
                    params = split_dbrw_params(p_match.group(1))
                    if not params:
                        continue
                        
                    raw_cube_param = params[0].strip()
                    matched_cube_key = None
                    
                    if raw_cube_param in cube_cache:
                        matched_cube_key = cube_cache[raw_cube_param]
                    else:
                        clean_param = raw_cube_param.strip('"').strip("'")
                        
                        if re.match(r'^(\'?[^\']+\'?!)?\$?[A-Z]+\$?\d+$', clean_param, re.IGNORECASE):
                            target_tab = tab_name
                            target_ref = clean_param
                            
                            if '!' in clean_param:
                                parts = clean_param.split('!')
                                target_tab = parts[0].strip("'")
                                target_ref = parts[1]
                                
                            target_ref = target_ref.replace('$', '').upper()
                            
                            if target_tab in ram_sheets:
                                target_tree = ram_sheets[target_tab]
                                
                                for tgt_cell in target_tree.findall('.//s:c', NS):
                                    if tgt_cell.attrib.get('r') == target_ref:
                                        tgt_f = tgt_cell.find('s:f', NS)
                                        
                                        if tgt_f is not None and tgt_f.text and 'VIEW' in tgt_f.text.upper():
                                            view_match = re.search(r'(?:_xll\.)?VIEW\s*\((.*)\)', tgt_f.text, re.IGNORECASE)
                                            if view_match:
                                                view_params = split_dbrw_params(view_match.group(1))
                                                if view_params:
                                                    legacy_cube = view_params[0].strip('"').strip("'")
                                                    
                                                    temp_idx = 1
                                                    for r_key, r_data in ram_mappings.items():
                                                        if r_key.lower() in legacy_cube.lower():
                                                            temp_idx = r_data['account_index']
                                                            matched_cube_key = r_key
                                                            break
                                                            
                                                    if matched_cube_key:
                                                        for find_str, replace_str in ram_mappings[matched_cube_key]['tab2_context'].items():
                                                            if find_str in view_params[0]:
                                                                view_params[0] = view_params[0].replace(find_str, replace_str)
                                                    
                                                    insert_pos = temp_idx
                                                    if insert_pos <= len(view_params):
                                                        view_params.insert(insert_pos, '"!"')
                                                    else:
                                                        view_params.append('"!"')
                                                    
                                                    tgt_f.text = f'VIEW({",".join(view_params)})'
                                                    tgt_v = tgt_cell.find('s:v', NS)
                                                    if tgt_v is not None:
                                                        tgt_cell.remove(tgt_v)
                                                        
                                                    cube_cache[raw_cube_param] = matched_cube_key
                                                    if target_tab != tab_name:
                                                        external_tab_cache[f"{target_tab}!{target_ref}"] = True

                                        else:
                                            val = get_cell_value(target_tree, target_ref, shared_strings)
                                            if val:
                                                for r_key in ram_mappings.keys():
                                                    if r_key.lower() in val.lower():
                                                        matched_cube_key = r_key
                                                        break
                                                cube_cache[raw_cube_param] = matched_cube_key
                                                
                                        break
                                        
                        else:
                            for r_key in ram_mappings.keys():
                                if r_key.lower() in clean_param.lower():
                                    matched_cube_key = r_key
                                    break
                            cube_cache[raw_cube_param] = matched_cube_key

                    if not matched_cube_key or matched_cube_key not in ram_mappings:
                        continue
                        
                    mapping_data = ram_mappings[matched_cube_key]
                    tab1_accounts = mapping_data['tab1_accounts']
                    tab_exceptions = mapping_data.get('tab_exceptions', {})
                    tab2_context = mapping_data['tab2_context']
                    tab3_blueprint = mapping_data['tab3_blueprint']
                    account_idx = mapping_data['account_index']
                    dz_headers = mapping_data['dropzone_headers']
                    
                    dz_col_map = {header: get_col_letter(start_dz_idx + i) for i, header in enumerate(dz_headers)}
                    audit_log[tab_name]['DZ_Cols'] = ", ".join(dz_col_map.values())
                    
                    mapped_params = list(params)
                    account_param_raw = None
                    
                    for p_idx, raw_param in enumerate(params):
                        if p_idx == 0:
                            continue
                            
                        if p_idx == account_idx - 1:
                            account_param_raw = raw_param
                            continue
                            
                        clean_param = raw_param.strip('"').strip("'")
                        
                        if re.match(r'^(\'?[^\']+\'?!)?\$?[A-Z]+\$?\d+$', clean_param, re.IGNORECASE):
                            target_tab = tab_name
                            target_ref = clean_param
                            
                            if '!' in clean_param:
                                parts = clean_param.split('!')
                                target_tab = parts[0].strip("'")
                                target_ref = parts[1]
                                
                            target_ref = target_ref.replace('$', '').upper()
                            cache_key = f"{target_tab}!{target_ref}"
                            
                            if cache_key not in header_cache:
                                if target_tab in ram_sheets:
                                    target_tree = ram_sheets[target_tab]
                                    
                                    for tgt_cell in target_tree.findall('.//s:c', NS):
                                        if tgt_cell.attrib.get('r') == target_ref:
                                            tgt_f = tgt_cell.find('s:f', NS)
                                            
                                            if tgt_f is not None and tgt_f.text and 'SUBNM' in tgt_f.text.upper():
                                                subnm_match = re.search(r'(?:_xll\.)?SUBNM\s*\((.*)\)', tgt_f.text, re.IGNORECASE)
                                                if subnm_match:
                                                    subnm_params = split_dbrw_params(subnm_match.group(1))
                                                    
                                                    for s_idx, s_val in enumerate(subnm_params):
                                                        new_s_val = s_val
                                                        for find_str, replace_str in tab2_context.items():
                                                            if find_str in new_s_val:
                                                                new_s_val = new_s_val.replace(find_str, replace_str)
                                                        subnm_params[s_idx] = new_s_val
                                                        
                                                    tgt_f.text = f'SUBNM({",".join(subnm_params)})'
                                                    tgt_v = tgt_cell.find('s:v', NS)
                                                    if tgt_v is not None:
                                                        tgt_cell.remove(tgt_v)
                                                        
                                            elif tgt_f is None:
                                                val = get_cell_value(target_tree, target_ref, shared_strings)
                                                if val and isinstance(val, str):
                                                    new_val = val
                                                    for find_str, replace_str in tab2_context.items():
                                                        if find_str in new_val:
                                                            new_val = new_val.replace(find_str, replace_str)
                                                            
                                                    if new_val != val:
                                                        tgt_cell.attrib['t'] = 'inlineStr'
                                                        for child in list(tgt_cell): tgt_cell.remove(child)
                                                        ET.SubElement(ET.SubElement(tgt_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = new_val
                                                        
                                            header_cache[cache_key] = True
                                            break
                                            
                        else:
                            new_param = raw_param
                            for find_str, replace_str in tab2_context.items():
                                if find_str in new_param:
                                    new_param = new_param.replace(find_str, replace_str)
                            mapped_params[p_idx] = new_param

                    base_dz_values = [""] * len(dz_headers)
                    active_clones = []
                    
                    if account_param_raw:
                        clean_acct_param = account_param_raw.strip('"').strip("'")
                        
                        if re.match(r'^(\'?[^\']+\'?!)?\$?[A-Z]+\$?\d+$', clean_acct_param, re.IGNORECASE):
                            acct_target_tab = tab_name
                            acct_target_ref = clean_acct_param
                            
                            if '!' in clean_acct_param:
                                parts = clean_acct_param.split('!')
                                acct_target_tab = parts[0].strip("'")
                                acct_target_ref = parts[1]
                                
                            acct_target_ref = acct_target_ref.replace('$', '').upper()
                            acct_cache_key = f"{acct_target_tab}!{acct_target_ref}"
                            acct_col_let = parse_col_letter(acct_target_ref)
                            
                            audit_log[tab_name]['Acct_Col'] = acct_col_let
                            
                            if acct_cache_key not in account_cache:
                                account_val = None
                                acct_cell_elem = None
                                
                                for c in row.findall('s:c', NS):
                                    if c.attrib.get('r') == acct_target_ref:
                                        acct_cell_elem = c
                                        account_val = get_cell_value(sheet_tree, acct_target_ref, shared_strings)
                                        break
                                        
                                normalized_account_val = normalize_account_name(account_val) if account_val else None
                                is_exception = (normalized_account_val in tab_exceptions) if normalized_account_val else False
                                
                                base_row_snapshot = None
                                if is_exception:
                                    base_row_snapshot = copy.deepcopy(row)
                                    
                                if account_val:
                                    audit_log[tab_name]['Evaluated'] += 1
                                    is_scenario_1 = False
                                    
                                    if is_exception:
                                        audit_log[tab_name]['Exceptions'] += 1
                                        is_scenario_1 = True
                                        exception_list = tab_exceptions[normalized_account_val]
                                        
                                        wd_account_1, cat_1 = exception_list[0]
                                        wd_account = wd_account_1
                                        base_dz_values = [cat_1] + [""] * (len(dz_headers) - 1)
                                        
                                        combo = (wd_account, tuple(base_dz_values))
                                        if combo in seen_wd_combos:
                                            if seen_wd_combos[combo] != normalized_account_val:
                                                base_dz_values = ["[DUPLICATE]"] * len(dz_headers)
                                                is_scenario_1 = False 
                                                audit_log[tab_name]['Duplicates'] += 1
                                        else:
                                            seen_wd_combos[combo] = normalized_account_val
                                            
                                        for split_wd, split_cat in exception_list[1:]:
                                            max_row += 1
                                            new_row_num = str(max_row)
                                            cloned_row = copy.deepcopy(base_row_snapshot)
                                            cloned_row.attrib['r'] = new_row_num
                                            
                                            for c in cloned_row.findall('s:c', NS):
                                                old_ref = c.attrib.get('r', '')
                                                if old_ref:
                                                    c_col_let = parse_col_letter(old_ref)
                                                    c.attrib['r'] = f"{c_col_let}{new_row_num}"
                                                    
                                            clone_acct_ref = f"{acct_col_let}{new_row_num}"
                                            clone_acct_cell = None
                                            for c in cloned_row.findall('s:c', NS):
                                                if c.attrib.get('r') == clone_acct_ref:
                                                    clone_acct_cell = c
                                                    break
                                            if not clone_acct_cell:
                                                clone_acct_cell = ET.Element(f'{{{NS_URL}}}c', {'r': clone_acct_ref})
                                                cloned_row.append(clone_acct_cell)
                                                
                                            clone_acct_cell.attrib['t'] = 'inlineStr'
                                            for child in list(clone_acct_cell): clone_acct_cell.remove(child)
                                            ET.SubElement(ET.SubElement(clone_acct_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str(split_wd)
                                            
                                            dz_values_clone = [split_cat] + [""] * (len(dz_headers) - 1)
                                            combo_clone = (split_wd, tuple(dz_values_clone))
                                            if combo_clone in seen_wd_combos:
                                                if seen_wd_combos[combo_clone] != normalized_account_val:
                                                    dz_values_clone = ["[DUPLICATE]"] * len(dz_headers)
                                                    audit_log[tab_name]['Duplicates'] += 1
                                            else:
                                                seen_wd_combos[combo_clone] = normalized_account_val
                                                
                                            for h_idx, dz_val in enumerate(dz_values_clone):
                                                target_dz_col = dz_col_map[dz_headers[h_idx]]
                                                dz_cell_ref = f"{target_dz_col}{new_row_num}"
                                                
                                                for c in cloned_row.findall('s:c', NS):
                                                    if c.attrib.get('r') == dz_cell_ref:
                                                        cloned_row.remove(c)
                                                        
                                                dz_cell = ET.Element(f'{{{NS_URL}}}c', {'r': dz_cell_ref, 't': 'inlineStr'})
                                                ET.SubElement(ET.SubElement(dz_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str(dz_val)
                                                cloned_row.append(dz_cell)
                                                
                                            vba_col_let = get_col_letter(vba_col_idx)
                                            vba_cell_ref = f"{vba_col_let}{new_row_num}"
                                            vba_cell = ET.Element(f'{{{NS_URL}}}c', {'r': vba_cell_ref, 't': 'inlineStr'})
                                            ET.SubElement(ET.SubElement(vba_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = f"[INSERT_BELOW_{row_num}]"
                                            cloned_row.append(vba_cell)
                                            
                                            for c_cell in cloned_row.findall('s:c', NS):
                                                c_f_tag = c_cell.find('s:f', NS)
                                                if c_f_tag is not None:
                                                    if not (c_f_tag.text and 'DBRW' in c_f_tag.text.upper()):
                                                        c_cell.remove(c_f_tag)
                                                        c_v_tag = c_cell.find('s:v', NS)
                                                        if c_v_tag is not None:
                                                            c_cell.remove(c_v_tag)

                                            cells_c = [c for c in cloned_row if c.tag == f'{{{NS_URL}}}c']
                                            cells_c.sort(key=lambda c: col_letter_to_idx(parse_col_letter(c.attrib.get('r', 'A'))))
                                            for c in cells_c: cloned_row.remove(c)
                                            cloned_row.extend(cells_c)
                                            
                                            cloned_rows_buffer.append(cloned_row)
                                            active_clones.append((new_row_num, cloned_row))
                                            
                                    elif normalized_account_val in tab1_accounts:
                                        wd_account, dz_values_raw = tab1_accounts[normalized_account_val]
                                        
                                        if wd_account is None or str(wd_account).strip() == "":
                                            base_dz_values = ["unused"] * len(dz_headers)
                                            audit_log[tab_name]['Incomplete'] += 1
                                        elif str(wd_account).strip().lower() == "exception" or any(str(v).strip().lower() == "exception" for v in dz_values_raw if v is not None):
                                            base_dz_values = [str(v) if v is not None else "" for v in dz_values_raw]
                                            audit_log[tab_name]['Exceptions'] += 1
                                        else:
                                            is_scenario_1 = True
                                            base_dz_values = [str(v) if v is not None else "" for v in dz_values_raw]
                                            
                                            combo = (wd_account, tuple(base_dz_values))
                                            if combo in seen_wd_combos:
                                                if seen_wd_combos[combo] != normalized_account_val:
                                                    base_dz_values = ["[DUPLICATE]"] * len(dz_headers)
                                                    is_scenario_1 = False 
                                                    audit_log[tab_name]['Duplicates'] += 1
                                                else:
                                                    audit_log[tab_name]['Converted'] += 1
                                            else:
                                                seen_wd_combos[combo] = normalized_account_val
                                                audit_log[tab_name]['Converted'] += 1
                                    else:
                                        base_dz_values = ["Unmapped"] * len(dz_headers)
                                        audit_log[tab_name]['Unmapped'] += 1

                                    if is_scenario_1 and acct_cell_elem is not None:
                                        acct_cell_elem.attrib['t'] = 'inlineStr'
                                        for child in list(acct_cell_elem):
                                            acct_cell_elem.remove(child)
                                        ET.SubElement(ET.SubElement(acct_cell_elem, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str(wd_account)

                                    for h_idx, dz_val in enumerate(base_dz_values):
                                        target_dz_col = dz_col_map[dz_headers[h_idx]]
                                        dz_cell_ref = f"{target_dz_col}{row_num}"
                                        
                                        for existing_c in row.findall('s:c', NS):
                                            if existing_c.attrib.get('r') == dz_cell_ref:
                                                row.remove(existing_c)
                                                break
                                                
                                        dz_cell = ET.Element(f'{{{NS_URL}}}c', {'r': dz_cell_ref, 't': 'inlineStr'})
                                        ET.SubElement(ET.SubElement(dz_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str(dz_val)
                                        row.append(dz_cell)

                                    cells = [c for c in row if c.tag == f'{{{NS_URL}}}c']
                                    cells.sort(key=lambda c: col_letter_to_idx(parse_col_letter(c.attrib.get('r', 'A'))))
                                    for c in cells: row.remove(c)
                                    row.extend(cells)

                                account_cache[acct_cache_key] = active_clones

                            else:
                                active_clones = account_cache[acct_cache_key]

                        else:
                            normalized_hardcode = normalize_account_name(clean_acct_param)
                            if normalized_hardcode in tab1_accounts:
                                mapped_params[account_idx - 1] = f'"{tab1_accounts[normalized_hardcode][0]}"'
                            elif normalized_hardcode in tab_exceptions:
                                mapped_params[account_idx - 1] = f'"{tab_exceptions[normalized_hardcode][0][0]}"'

                    primary_cube_ref = raw_cube_param.replace('$', '').upper()
                    
                    def base_replacer(match):
                        full_match = match.group(0)
                        inner_m = re.search(DBRW_EXTRACT_RGX, full_match, re.IGNORECASE)
                        if not inner_m: return full_match
                        
                        local_params = split_dbrw_params(inner_m.group(1))
                        if not local_params: return full_match
                        
                        cur_cube = local_params[0].strip().replace('$', '').upper()
                        if cur_cube != primary_cube_ref:
                            return full_match 
                        
                        local_bp = tab3_blueprint.lstrip('=')
                        for idx, p_val in enumerate(local_params):
                            local_bp = local_bp.replace(f"{{Param_{idx}}}", p_val)
                            
                        for header_name, col_let in dz_col_map.items():
                            local_bp = local_bp.replace(f"{{{header_name}}}", f"${col_let}{row_num}")
                            
                        for old_val, new_val in tab2_context.items():
                            local_bp = local_bp.replace(old_val, new_val)
                            
                        return local_bp

                    updated_formula = re.sub(DBRW_SUB_RGX, base_replacer, old_formula, flags=re.IGNORECASE)

                    si_val = f_tag.attrib.get('si')
                    if si_val is not None:
                        popped_si_values.add(si_val)

                    f_tag.text = updated_formula.lstrip('=')
                    f_tag.attrib.pop('t', None)
                    f_tag.attrib.pop('si', None)
                    f_tag.attrib.pop('ref', None)

                    v_tag = cell.find('s:v', NS)
                    if v_tag is not None:
                        cell.remove(v_tag)

                    for new_row_num, cloned_row in active_clones:
                        for c_cell in cloned_row.findall('s:c', NS):
                            if c_cell.attrib.get('r') == cell.attrib.get('r').replace(row_num, new_row_num):
                                c_f_tag = c_cell.find('s:f', NS)
                                if c_f_tag is not None and c_f_tag.text and 'DBRW' in c_f_tag.text.upper():
                                    c_old_formula = c_f_tag.text
                                    
                                    def clone_replacer(match):
                                        full_match = match.group(0)
                                        inner_m = re.search(DBRW_EXTRACT_RGX, full_match, re.IGNORECASE)
                                        if not inner_m: return full_match
                                        
                                        local_params = split_dbrw_params(inner_m.group(1))
                                        if not local_params: return full_match
                                        
                                        cur_cube = local_params[0].strip().replace('$', '').upper()
                                        if cur_cube != primary_cube_ref:
                                            return full_match
                                        
                                        local_bp = tab3_blueprint.lstrip('=')
                                        for idx, p_val in enumerate(local_params):
                                            p_val_renum = re.sub(r'(\$?[A-Z]+\$?)' + str(row_num) + r'\b', r'\g<1>' + str(new_row_num), p_val, flags=re.IGNORECASE)
                                            local_bp = local_bp.replace(f"{{Param_{idx}}}", p_val_renum)
                                            
                                        for header_name, col_let in dz_col_map.items():
                                            local_bp = local_bp.replace(f"{{{header_name}}}", f"${col_let}{new_row_num}")
                                            
                                        for old_val, new_val in tab2_context.items():
                                            local_bp = local_bp.replace(old_val, new_val)
                                            
                                        return local_bp

                                    c_updated = re.sub(DBRW_SUB_RGX, clone_replacer, c_old_formula, flags=re.IGNORECASE)
                                    
                                    c_si_val = c_f_tag.attrib.get('si')
                                    if c_si_val is not None:
                                        popped_si_values.add(c_si_val)
                                        
                                    c_f_tag.text = c_updated.lstrip('=')
                                    c_f_tag.attrib.pop('t', None)
                                    c_f_tag.attrib.pop('si', None)
                                    c_f_tag.attrib.pop('ref', None)
                                    
                                    c_v_tag = c_cell.find('s:v', NS)
                                    if c_v_tag is not None:
                                        c_cell.remove(c_v_tag)

        for clone in cloned_rows_buffer:
            sheet_tree.append(clone)

        if popped_si_values:
            for cell in sheet_tree.findall('.//s:c', NS):
                f_tag = cell.find('s:f', NS)
                if f_tag is not None and f_tag.attrib.get('t') == 'shared':
                    child_si = f_tag.attrib.get('si')
                    if child_si in popped_si_values and f_tag.text is None:
                        cell.remove(f_tag)

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

    # ==========================================
    # PHASE 8 - CLEAN AUDIT TAB INJECTION (VBA HANDSHAKE)
    # ==========================================
    if structural_files['xl/workbook.xml'] and audit_log:
        try:
            def escape_xml(text):
                s = str(text)
                s = s.replace("&amp;", "&") 
                return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")

            # 1. Parse workbook.xml to register the new sheet
            wb_data = structural_files['xl/workbook.xml'].decode('utf-8')
            sheet_ids = [int(m) for m in re.findall(r'sheetId="(\d+)"', wb_data)]
            new_sheet_id = max(sheet_ids) + 1 if sheet_ids else 99
            
            # 2. Parse workbook.xml.rels to assign a fresh Relationship ID
            rels_data = structural_files['xl/_rels/workbook.xml.rels'].decode('utf-8')
            r_ids = [int(m) for m in re.findall(r'rId(\d+)"', rels_data)]
            max_r_id = max(r_ids) if r_ids else 99
            new_r_id = f"rId{max_r_id + 1}"

            # 3. Inject new sheet and relationship nodes
            new_sheet_node = f'<sheet name="Conversion Audit" sheetId="{new_sheet_id}" r:id="{new_r_id}"/>'
            wb_data = wb_data.replace('</sheets>', f'{new_sheet_node}</sheets>')
            
            new_rel_node = f'<Relationship Id="{new_r_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet_audit_{new_sheet_id}.xml"/>'
            rels_data = rels_data.replace('</Relationships>', f'{new_rel_node}</Relationships>')

            # 4. Inject Content Types Override
            ct_data = structural_files['[Content_Types].xml'].decode('utf-8')
            new_override = f'<Override PartName="/xl/worksheets/sheet_audit_{new_sheet_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            ct_data = ct_data.replace('</Types>', f'{new_override}</Types>')

            # Write updated structural XMLs back to the zip
            output_zip.writestr('xl/workbook.xml', wb_data.encode('utf-8'))
            output_zip.writestr('xl/_rels/workbook.xml.rels', rels_data.encode('utf-8'))
            output_zip.writestr('[Content_Types].xml', ct_data.encode('utf-8'))
            if structural_files.get('docProps/app.xml'):
                output_zip.writestr('docProps/app.xml', structural_files['docProps/app.xml'])

            # 5. Build raw Audit Worksheet XML
            audit_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            audit_xml += f'<worksheet xmlns="{NS_URL}">\n'
            audit_xml += '<sheetData>\n'
            
            headers = ['Sheet Name', 'Account Column', 'Dropzone Column(s)', 'Total Evaluated', 'Converted (Success)', 'Exceptions', 'Incomplete Match', 'Unmapped (Orphans)', 'Duplicates Suppressed']
            audit_xml += '<row r="1">\n'
            for col_idx, header_val in enumerate(headers, start=1):
                col_let = get_col_letter(col_idx)
                safe_header = escape_xml(header_val)
                audit_xml += f'<c r="{col_let}1" t="inlineStr"><is><t>{safe_header}</t></is></c>\n'
            audit_xml += '</row>\n'

            row_num = 2
            for tab_n, stats in audit_log.items():
                audit_xml += f'<row r="{row_num}">\n'
                
                # Column A: Sheet Name
                col_let = get_col_letter(1)
                safe_tab_name = escape_xml(tab_n)
                audit_xml += f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{safe_tab_name}</t></is></c>\n'
                
                # Column B: Account Column
                col_let = get_col_letter(2)
                audit_xml += f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{escape_xml(stats.get("Acct_Col", "N/A"))}</t></is></c>\n'
                
                # Column C: Dropzone Column(s)
                col_let = get_col_letter(3)
                audit_xml += f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{escape_xml(stats.get("DZ_Cols", "N/A"))}</t></is></c>\n'
                
                # Columns D-I: Numerical Conversion Stats
                numeric_stats = [
                    stats['Evaluated'],
                    stats['Converted'],
                    stats['Exceptions'],
                    stats['Incomplete'],
                    stats['Unmapped'],
                    stats['Duplicates']
                ]
                
                for i, num_val in enumerate(numeric_stats, start=4):
                    col_let = get_col_letter(i)
                    audit_xml += f'<c r="{col_let}{row_num}"><v>{num_val}</v></c>\n'

                audit_xml += '</row>\n'
                row_num += 1

            audit_xml += '</sheetData>\n</worksheet>'
            
            # Write new sheet file
            output_zip.writestr(f'xl/worksheets/sheet_audit_{new_sheet_id}.xml', audit_xml.encode('utf-8'))
            print(f"\n  ├─ Successfully Created & Injected 'Conversion Audit' tab ({len(audit_log)} sheets logged).")

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