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

# Optimized: Instant Value Lookup from TCell
def get_cell_value(tcell, shared_strings):
    if tcell is None:
        return None
    cell_type = tcell.attrib.get('t')
    if cell_type == 'inlineStr':
        is_tag = tcell.find('s:is/s:t', NS)
        val = is_tag.text.strip() if is_tag is not None and is_tag.text else None
        return val.lstrip("'").lstrip("`") if val else None
    v_tag = tcell.find('s:v', NS)
    if v_tag is not None and v_tag.text:
        if cell_type == 's':
            idx = int(v_tag.text)
            val = shared_strings[idx].strip() if idx < len(shared_strings) else v_tag.text.strip()
        else:
            val = v_tag.text.strip()
        return val.lstrip("'").lstrip("`") if val else None
    return None

def apply_context_replacements(text, context_dict):
    if not text or not isinstance(text, str): 
        return text
    if not context_dict:
        return text
        
    sorted_keys = sorted(context_dict.keys(), key=len, reverse=True)
    
    pattern_parts = []
    for key in sorted_keys:
        esc_k = re.escape(key)
        pattern_parts.append(rf"(?<![A-Za-z0-9_]){esc_k}(?![A-Za-z0-9_])")
        
    pattern = re.compile('|'.join(pattern_parts), re.IGNORECASE)
    lower_map = {k.lower(): v for k, v in context_dict.items()}
    
    def replacer(match):
        return lower_map[match.group(0).lower()]
        
    return pattern.sub(replacer, text)

# Optimized: Instant Target Resolution via Dictionary Map
def resolve_target_cell(start_tab, start_ref, ram_sheets, ram_cell_maps):
    cur_tab = start_tab
    cur_ref = start_ref.replace('$', '').upper()
    visited = set()
    
    while True:
        cache_k = f"{cur_tab}!{cur_ref}"
        if cache_k in visited:
            break
        visited.add(cache_k)
        
        if cur_tab not in ram_sheets:
            return cur_tab, cur_ref, None, None
            
        tcell = ram_cell_maps[cur_tab].get(cur_ref)
        if tcell is None:
            return cur_tab, cur_ref, None, None
            
        tf = tcell.find('s:f', NS)
        if tf is not None and tf.text:
            f_txt = tf.text.strip()
            if any(x in f_txt.upper() for x in ['SUBNM', 'VIEW', 'DBRW', 'TM1RPTVIEW']):
                return cur_tab, cur_ref, tcell, tf
            
            link_m = re.match(r'^=?\s*\+?\s*(\'?[^\']+\'?!)?\$?([A-Z]+)\$?(\d+)$', f_txt, re.IGNORECASE)
            if link_m:
                next_tab_str = link_m.group(1)
                if next_tab_str:
                    cur_tab = next_tab_str.strip('!').strip("'")
                cur_ref = f"{link_m.group(2).upper()}{link_m.group(3)}"
                continue
            else:
                return cur_tab, cur_ref, tcell, tf
        else:
            return cur_tab, cur_ref, tcell, None
            
    return cur_tab, cur_ref, None, None

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
    ram_cell_maps = {} 
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
            continue

        header_text = raw_xml_str[:sheet_data_match.start(2)]
        inner_text = sheet_data_match.group(2)
        footer_text = raw_xml_str[sheet_data_match.end(2):]

        header_text = re.sub(r'<dimension[^>]*/>', '', header_text)
        header_text = re.sub(r'<dimension[^>]*>.*?</dimension>', '', header_text)

        ns_declarations = " ".join(re.findall(r'xmlns(?::[a-zA-Z0-9\-]+)?="[^"]*"', header_text))
        dummy_xml = f'<root {ns_declarations}>{inner_text}</root>'
        
        try:
            sheet_tree = ET.fromstring(dummy_xml)
            ram_sheets[tab_name] = sheet_tree
            ram_cell_maps[tab_name] = {c.attrib.get('r'): c for c in sheet_tree.findall('.//s:c', NS) if c.attrib.get('r')}
            sheet_wrappers[tab_name] = (header_text, footer_text)
            sheet_path_map[tab_name] = sheet_path
            
            audit_log[tab_name] = {
                'Acct_Col': 'N/A', 'DZ_Cols': 'N/A',
                'Evaluated': 0, 'Converted': 0, 'Exceptions': 0, 'Incomplete': 0, 'Unmapped': 0, 'Duplicates': 0
            }
        except ET.ParseError:
            pass

    cube_cache = {}
    cube_cache_hits = {}   
    header_cache = {}
    header_cache_hits = {}  
    
    account_val_cache = {}
    account_updated_cache = set()
    
    modified_tabs = set()

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
            continue
            
        modified_tabs.add(tab_name)
        print(f"\n  ├─ [Processing Tab]: '{tab_name}' (Sheet {s_num})")

        # ==========================================
        # PRE-PROCESSING: FULLY FLATTEN SHARED FORMULAS
        # ==========================================
        for r_elem in sheet_tree.findall('.//s:row', NS):
            r_elem.attrib.pop('spans', None)

        shared_formulas = {}
        for c in sheet_tree.findall('.//s:c', NS):
            f_tag = c.find('s:f', NS)
            if f_tag is not None:
                si = f_tag.attrib.get('si')
                if si is not None and f_tag.text:
                    shared_formulas[si] = {
                        'text': f_tag.text, 
                        'master_row': parse_row_num(c.attrib.get('r', 'A1'))
                    }
                    
        for c in sheet_tree.findall('.//s:c', NS):
            f_tag = c.find('s:f', NS)
            if f_tag is not None and f_tag.attrib.get('t') == 'shared':
                if not f_tag.text:
                    si = f_tag.attrib.get('si')
                    if si in shared_formulas:
                        cur_row = parse_row_num(c.attrib.get('r', 'A1'))
                        shifted_text = re.sub(
                            r'(\$?[A-Za-z]+)' + str(shared_formulas[si]['master_row']) + r'\b', 
                            r'\g<1>' + str(cur_row), 
                            shared_formulas[si]['text'], 
                            flags=re.IGNORECASE
                        )
                        f_tag.text = shifted_text
                        
                        v_tag = c.find('s:v', NS)
                        if v_tag is not None:
                            c.remove(v_tag)
                        c.attrib.pop('t', None)
                    else:
                        c.remove(f_tag)
                        v_tag = c.find('s:v', NS)
                        if v_tag is not None:
                            c.remove(v_tag)
                        c.attrib.pop('t', None)
                        continue 
                
                f_tag.attrib.pop('t', None)
                f_tag.attrib.pop('si', None)
                f_tag.attrib.pop('ref', None)

        seen_wd_combos = {}
        popped_si_values = set()
        cloned_rows_buffer = []
        all_active_clones = []
        
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

        header_ownership = {}
        tab_string_reroutes = {}
        diverted_values_map = {}
        
        header_data_row_num = str(max_row + 1)
        header_data_row = ET.Element(f'{{{NS_URL}}}row', {'r': header_data_row_num})
        max_row += 1 
        
        header_conflict_col_idx = 2
        header_row_has_data = False

        for row in sheet_tree.findall('.//s:row', NS):
            row_num_str = row.attrib.get('r')
            if not row_num_str:
                continue
            row_num = row_num_str
            
            row_active_clones = {}
            
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
                        
                    # ==========================================
                    # PHASE 3 - CUBE RESOLUTION
                    # ==========================================
                    raw_cube_param = params[0].strip()
                    clean_param = raw_cube_param.strip('"').strip("'")
                    
                    is_cell_ref = bool(re.match(r'^(\'?[^\']+\'?!)?\$?[A-Z]+\$?\d+$', clean_param, re.IGNORECASE))
                    matched_cube_key = None
                    
                    if is_cell_ref:
                        parts = clean_param.split('!') if '!' in clean_param else [tab_name, clean_param]
                        start_tab = parts[0].strip("'") if len(parts) > 1 else tab_name
                        start_ref = parts[-1].replace('$', '').upper()
                        
                        final_tab, final_ref, tgt_cell, tgt_f = resolve_target_cell(start_tab, start_ref, ram_sheets, ram_cell_maps)
                        cache_key = f"{final_tab}!{final_ref}"

                        if cache_key in cube_cache:
                            cube_cache_hits[cache_key] = cube_cache_hits.get(cache_key, 0) + 1
                            matched_cube_key = cube_cache[cache_key]
                        else:
                            cube_cache_hits[cache_key] = 0
                            if tgt_cell is not None:
                                f_upper = tgt_f.text.upper() if tgt_f is not None and tgt_f.text else ""
                                
                                if 'TM1RPTVIEW' in f_upper:
                                    legacy_cube_match = re.search(r'TM1RPTVIEW\s*\(\s*["\']([^"\']+)["\']', tgt_f.text, re.IGNORECASE)
                                    if legacy_cube_match:
                                        legacy_cube = legacy_cube_match.group(1)
                                        for r_key in ram_mappings.keys():
                                            if r_key.lower() in legacy_cube.lower():
                                                matched_cube_key = r_key
                                                break
                                                
                                        if matched_cube_key:
                                            tgt_f.text = apply_context_replacements(tgt_f.text, ram_mappings[matched_cube_key]['tab2_context'])
                                            tgt_v = tgt_cell.find('s:v', NS)
                                            if tgt_v is not None:
                                                tgt_cell.remove(tgt_v)
                                            tgt_cell.attrib.pop('t', None)  
                                                
                                            cube_cache[cache_key] = matched_cube_key
                                            modified_tabs.add(final_tab)

                                elif 'VIEW' in f_upper:
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
                                                view_params[0] = apply_context_replacements(view_params[0], ram_mappings[matched_cube_key]['tab2_context'])
                                            
                                            insert_pos = temp_idx
                                            if insert_pos <= len(view_params):
                                                view_params.insert(insert_pos, '"!"')
                                            else:
                                                view_params.append('"!"')
                                            
                                            tgt_f.text = f'VIEW({",".join(view_params)})'
                                            tgt_v = tgt_cell.find('s:v', NS)
                                            if tgt_v is not None:
                                                tgt_cell.remove(tgt_v)
                                            tgt_cell.attrib.pop('t', None)  
                                                
                                            cube_cache[cache_key] = matched_cube_key
                                            modified_tabs.add(final_tab)

                                else:
                                    val = get_cell_value(tgt_cell, shared_strings)
                                    if val:
                                        for r_key in ram_mappings.keys():
                                            if r_key.lower() in val.lower():
                                                matched_cube_key = r_key
                                                break
                                        if matched_cube_key:
                                            new_val = apply_context_replacements(val, ram_mappings[matched_cube_key]['tab2_context'])
                                            if new_val != val:
                                                tgt_cell.attrib['t'] = 'inlineStr'
                                                for child in list(tgt_cell): tgt_cell.remove(child)
                                                ET.SubElement(ET.SubElement(tgt_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = new_val
                                                modified_tabs.add(final_tab)
                                        cube_cache[cache_key] = matched_cube_key
                    else:
                        cache_key = raw_cube_param
                        if cache_key in cube_cache:
                            cube_cache_hits[cache_key] = cube_cache_hits.get(cache_key, 0) + 1
                            matched_cube_key = cube_cache[cache_key]
                        else:
                            for r_key in ram_mappings.keys():
                                if r_key.lower() in clean_param.lower():
                                    matched_cube_key = r_key
                                    break
                            cube_cache[cache_key] = matched_cube_key

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
                    
                    # ==========================================
                    # PHASE 4 - HEADER RESOLUTION & CACHE UPDATE
                    # ==========================================
                    for p_idx, raw_param in enumerate(params):
                        if p_idx == 0:
                            continue
                        if p_idx == account_idx - 1:
                            account_param_raw = raw_param
                            continue
                            
                        clean_param = raw_param.strip('"').strip("'")
                        is_cell_ref = bool(re.match(r'^(\'?[^\']+\'?!)?\$?[A-Z]+\$?\d+$', clean_param, re.IGNORECASE))
                        
                        if is_cell_ref:
                            parts = clean_param.split('!') if '!' in clean_param else [tab_name, clean_param]
                            start_tab = parts[0].strip("'") if len(parts) > 1 else tab_name
                            start_ref = parts[-1].replace('$', '').upper()
                            
                            final_tab, final_ref, tgt_cell, tgt_f = resolve_target_cell(start_tab, start_ref, ram_sheets, ram_cell_maps)
                            cache_key = f"{final_tab}!{final_ref}"
                            
                            raw_orig_text = None
                            is_subnm = False
                            val = None 
                            
                            if tgt_f is not None and tgt_f.text and 'SUBNM' in tgt_f.text.upper():
                                is_subnm = True
                                raw_orig_text = "SUBNM"
                            elif tgt_cell is not None:
                                val = get_cell_value(tgt_cell, shared_strings)
                                raw_orig_text = val if val and isinstance(val, str) else str(val) if val else ""
                                
                            simulated_new_val = raw_orig_text
                            if not is_subnm and raw_orig_text:
                                simulated_new_val = apply_context_replacements(raw_orig_text, tab2_context)
                                
                            divert_conflict = False
                            
                            if cache_key in header_ownership:
                                owning_cube = header_ownership[cache_key]['cube']
                                written_val = header_ownership[cache_key]['written_val']
                                
                                if owning_cube != matched_cube_key and written_val != simulated_new_val:
                                    divert_conflict = True
                            else:
                                header_ownership[cache_key] = {'cube': matched_cube_key, 'written_val': simulated_new_val}
                            
                            if divert_conflict:
                                if not header_row_has_data:
                                    lowest_cube_row = float('inf')
                                    for k in cube_cache.keys():
                                        if k.startswith(f"{tab_name}!"):
                                            ref = k.split('!')[1]
                                            r_num = int(parse_row_num(ref))
                                            if r_num < lowest_cube_row: lowest_cube_row = r_num
                                    if lowest_cube_row == float('inf'): lowest_cube_row = 1
                                    
                                    title_ref = f"A{header_data_row_num}"
                                    title_cell = ET.Element(f'{{{NS_URL}}}c', {'r': title_ref, 't': 'inlineStr'})
                                    ET.SubElement(ET.SubElement(title_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = "Header Data"
                                    header_data_row.append(title_cell)
                                    
                                    vba_ref = f"{get_col_letter(vba_col_idx)}{header_data_row_num}"
                                    vba_cell = ET.Element(f'{{{NS_URL}}}c', {'r': vba_ref, 't': 'inlineStr'})
                                    ET.SubElement(ET.SubElement(vba_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = f"[INSERT_ABOVE_{lowest_cube_row}]"
                                    header_data_row.append(vba_cell)
                                    
                                    header_row_has_data = True
                                    
                                str_val = str(simulated_new_val).strip() if simulated_new_val is not None else ""
                                str_key = str_val.upper()
                                
                                if str_key in diverted_values_map:
                                    divert_ref_abs = diverted_values_map[str_key]
                                    tab_string_reroutes[raw_param] = divert_ref_abs
                                else:
                                    divert_col_let = get_col_letter(header_conflict_col_idx)
                                    divert_ref = f"{divert_col_let}{header_data_row_num}"
                                    divert_ref_abs = f"${divert_col_let}${header_data_row_num}"
                                    
                                    divert_cell = ET.Element(f'{{{NS_URL}}}c', {'r': divert_ref, 't': 'inlineStr'})
                                    ET.SubElement(ET.SubElement(divert_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str_val
                                    header_data_row.append(divert_cell)
                                    
                                    diverted_values_map[str_key] = divert_ref_abs
                                    tab_string_reroutes[raw_param] = divert_ref_abs
                                    header_conflict_col_idx += 1
                            else:
                                if cache_key not in header_cache:
                                    header_cache_hits[cache_key] = 0
                                    if tgt_cell is not None:
                                        if is_subnm:
                                            subnm_match = re.search(r'(?:_xll\.)?SUBNM\s*\((.*)\)', tgt_f.text, re.IGNORECASE)
                                            extracted_element = "SUBNM" # Default fallback
                                            
                                            if subnm_match:
                                                subnm_params = split_dbrw_params(subnm_match.group(1))
                                                
                                                # Extract the 3rd parameter (Index 2) before we lose the original context
                                                if len(subnm_params) > 2:
                                                    extracted_element = subnm_params[2].replace('"', '').replace("'", "").strip()
                                                    
                                                for s_idx, s_val in enumerate(subnm_params):
                                                    subnm_params[s_idx] = apply_context_replacements(s_val, tab2_context)
                                                tgt_f.text = f'SUBNM({",".join(subnm_params)})'
                                                tgt_v = tgt_cell.find('s:v', NS)
                                                if tgt_v is not None:
                                                    tgt_cell.remove(tgt_v)
                                                tgt_cell.attrib.pop('t', None)  
                                                modified_tabs.add(final_tab)
                                                
                                            # Write the actual element (e.g., "YTD") to the cache so the sniffer finds it
                                            header_cache[cache_key] = extracted_element
                                                
                                        else:
                                            if val and isinstance(val, str):
                                                if simulated_new_val != val:
                                                    tgt_cell.attrib['t'] = 'inlineStr'
                                                    for child in list(tgt_cell): tgt_cell.remove(child)
                                                    ET.SubElement(ET.SubElement(tgt_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = simulated_new_val
                                                    modified_tabs.add(final_tab)
                                            header_cache[cache_key] = val
                                    else:
                                        header_cache[cache_key] = None
                                else:
                                    header_cache_hits[cache_key] = header_cache_hits.get(cache_key, 0) + 1
                        else:
                            mapped_params[p_idx] = apply_context_replacements(raw_param, tab2_context)

                    # ==========================================
                    # PHASE 4.5 - PANDL YTD SNIFFER
                    # ==========================================
                    apply_ytd_rule = False
                    if matched_cube_key and 'PANDL' in matched_cube_key.upper():
                        last_param_raw = params[-1]
                        clean_last_param = last_param_raw.strip('"').strip("'")
                        
                        if re.match(r'^(\'?[^\']+\'?!)?\$?[A-Z]+\$?\d+$', clean_last_param, re.IGNORECASE):
                            parts = clean_last_param.split('!') if '!' in clean_last_param else [tab_name, clean_last_param]
                            start_tab = parts[0].strip("'") if len(parts) > 1 else tab_name
                            start_ref = parts[-1].replace('$', '').upper()
                            
                            final_tab, final_ref, _, _ = resolve_target_cell(start_tab, start_ref, ram_sheets, ram_cell_maps)
                            cache_key = f"{final_tab}!{final_ref}"
                            
                            orig_val = header_cache.get(cache_key)
                            orig_val_str = str(orig_val).strip().upper() if orig_val else ""
                            if orig_val_str in ["YTD", "YTD_CDN", "PTD_CDN"]:
                                apply_ytd_rule = True
                        elif clean_last_param.upper().replace("'", "").replace("`", "").strip() in ["YTD", "YTD_CDN", "PTD_CDN"]:
                            apply_ytd_rule = True

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
                            
                            if acct_cache_key not in row_active_clones:
                                if acct_cache_key not in account_val_cache:
                                    tgt_acct_cell = ram_cell_maps.get(acct_target_tab, {}).get(acct_target_ref)
                                    account_val = get_cell_value(tgt_acct_cell, shared_strings)
                                    account_val_cache[acct_cache_key] = account_val
                                else:
                                    account_val = account_val_cache[acct_cache_key]
                                
                                normalized_account_val = normalize_account_name(account_val) if account_val else None
                                is_exception = (normalized_account_val in tab_exceptions) if normalized_account_val else False
                                
                                base_row_snapshot = None
                                if is_exception:
                                    base_row_snapshot = copy.deepcopy(row)
                                
                                if account_val:
                                    audit_log[tab_name]['Evaluated'] += 1
                                    is_scenario_1 = False
                                    wd_account = None
                                    
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
                                                        
                                                if str(dz_val).strip() != "":
                                                    dz_cell = ET.Element(f'{{{NS_URL}}}c', {'r': dz_cell_ref, 't': 'inlineStr'})
                                                    ET.SubElement(ET.SubElement(dz_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str(dz_val)
                                                    cloned_row.append(dz_cell)
                                                
                                            vba_col_let = get_col_letter(vba_col_idx)
                                            vba_cell_ref = f"{vba_col_let}{new_row_num}"
                                            vba_cell = ET.Element(f'{{{NS_URL}}}c', {'r': vba_cell_ref, 't': 'inlineStr'})
                                            ET.SubElement(ET.SubElement(vba_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = f"[INSERT_BELOW_{row_num}]"
                                            cloned_row.append(vba_cell)
                                            
                                            cells_c = [c for c in cloned_row if c.tag == f'{{{NS_URL}}}c']
                                            cells_c.sort(key=lambda c: col_letter_to_idx(parse_col_letter(c.attrib.get('r', 'A'))))
                                            for c in cells_c: cloned_row.remove(c)
                                            cloned_row.extend(cells_c)
                                            
                                            cloned_rows_buffer.append(cloned_row)
                                            active_clones.append((row_num, new_row_num, cloned_row))
                                            
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

                                    for h_idx, dz_val in enumerate(base_dz_values):
                                        target_dz_col = dz_col_map[dz_headers[h_idx]]
                                        dz_cell_ref = f"{target_dz_col}{row_num}"
                                        
                                        for existing_c in row.findall('s:c', NS):
                                            if existing_c.attrib.get('r') == dz_cell_ref:
                                                row.remove(existing_c)
                                                break
                                        
                                        if str(dz_val).strip() != "":
                                            dz_cell = ET.Element(f'{{{NS_URL}}}c', {'r': dz_cell_ref, 't': 'inlineStr'})
                                            ET.SubElement(ET.SubElement(dz_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str(dz_val)
                                            row.append(dz_cell)

                                    cells = [c for c in row if c.tag == f'{{{NS_URL}}}c']
                                    cells.sort(key=lambda c: col_letter_to_idx(parse_col_letter(c.attrib.get('r', 'A'))))
                                    for c in cells: row.remove(c)
                                    row.extend(cells)
                                    
                                    if is_scenario_1 and wd_account is not None:
                                        if acct_cache_key not in account_updated_cache:
                                            tgt_acct_cell = ram_cell_maps.get(acct_target_tab, {}).get(acct_target_ref)
                                            if tgt_acct_cell is not None:
                                                tgt_acct_cell.attrib['t'] = 'inlineStr'
                                                for child in list(tgt_acct_cell): tgt_acct_cell.remove(child)
                                                ET.SubElement(ET.SubElement(tgt_acct_cell, f'{{{NS_URL}}}is'), f'{{{NS_URL}}}t').text = str(wd_account)
                                                modified_tabs.add(acct_target_tab)
                                            account_updated_cache.add(acct_cache_key)

                                row_active_clones[acct_cache_key] = active_clones
                            else:
                                active_clones = row_active_clones[acct_cache_key]
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
                            if p_val in tab_string_reroutes:
                                p_val = tab_string_reroutes[p_val]
                                
                            if apply_ytd_rule and idx == 3:
                                p_val = f'{p_val}&"@@_SFX_99_@@"'
                            local_bp = local_bp.replace(f"{{Param_{idx}}}", p_val)
                            
                        for header_name, col_let in dz_col_map.items():
                            local_bp = local_bp.replace(f"{{{header_name}}}", f"${col_let}{row_num}")
                            
                        local_bp = apply_context_replacements(local_bp, tab2_context)
                        local_bp = local_bp.replace("@@_SFX_99_@@", "YTD")
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
                    cell.attrib.pop('t', None)

            for active_clones_list in row_active_clones.values():
                all_active_clones.extend(active_clones_list)

        # ==========================================
        # PHASE 9 - DELAYED CLONE EVALUATION
        # ==========================================
        for row_num, new_row_num, cloned_row in all_active_clones:
            for c_cell in cloned_row.findall('s:c', NS):
                c_f_tag = c_cell.find('s:f', NS)
                if c_f_tag is not None:
                    if c_f_tag.text:
                        c_f_tag.attrib.pop('t', None)
                        c_f_tag.attrib.pop('si', None)
                        c_f_tag.attrib.pop('ref', None)

                        shifted_formula = re.sub(
                            r'(\$?[A-Za-z]+)' + str(row_num) + r'\b', 
                            r'\g<1>' + str(new_row_num), 
                            c_f_tag.text, 
                            flags=re.IGNORECASE
                        )

                        if 'DBRW' in shifted_formula.upper():
                            def clone_universal_replacer(match):
                                full_match = match.group(0)
                                inner_m = re.search(DBRW_EXTRACT_RGX, full_match, re.IGNORECASE)
                                if not inner_m: return full_match
                                
                                local_params = split_dbrw_params(inner_m.group(1))
                                if not local_params: return full_match
                                
                                cur_cube_raw = local_params[0].strip().strip('"').strip("'")
                                cur_cube = cur_cube_raw.replace('$', '').upper()
                                
                                matched_ck = None
                                if "!" in cur_cube:
                                    parts = cur_cube.split("!")
                                    temp_tab = parts[0].strip("'")
                                    temp_ref = parts[1]
                                    cache_k = f"{temp_tab}!{temp_ref}"
                                else:
                                    cache_k = f"{tab_name}!{cur_cube}"
                                    
                                if cache_k in cube_cache:
                                    matched_ck = cube_cache[cache_k]
                                else:
                                    for r_key in ram_mappings.keys():
                                        if r_key.lower() in cur_cube_raw.lower():
                                            matched_ck = r_key
                                            break
                                            
                                if not matched_ck:
                                    return full_match
                                    
                                c_mapping = ram_mappings[matched_ck]
                                c_bp = c_mapping['tab3_blueprint'].lstrip('=')
                                c_ctx = c_mapping['tab2_context']
                                c_dz_headers = c_mapping['dropzone_headers']
                                c_dz_col_map = {header: get_col_letter(start_dz_idx + i) for i, header in enumerate(c_dz_headers)}
                                
                                is_ytd = False
                                if 'PANDL' in matched_ck.upper():
                                    last_p = local_params[-1].strip('"').strip("'")
                                    if re.match(r'^(\'?[^\']+\'?!)?\$?[A-Z]+\$?\d+$', last_p, re.IGNORECASE):
                                        parts = last_p.split('!') if '!' in last_p else [tab_name, last_p]
                                        temp_tab = parts[0].strip("'") if len(parts) > 1 else tab_name
                                        temp_ref = parts[-1].replace('$', '').upper()
                                        c_key = f"{temp_tab}!{temp_ref}"
                                        
                                        orig_val = header_cache.get(c_key)
                                        orig_val_str = str(orig_val).strip().upper() if orig_val else ""
                                        if orig_val_str in ["YTD", "YTD_CDN", "PTD_CDN"]:
                                            is_ytd = True
                                    elif last_p.upper().replace("'", "").replace("`", "").strip() in ["YTD", "YTD_CDN", "PTD_CDN"]:
                                        is_ytd = True
                                        
                                for idx, p_val in enumerate(local_params):
                                    if p_val in tab_string_reroutes:
                                        p_val = tab_string_reroutes[p_val]
                                        
                                    if is_ytd and idx == 3:
                                        p_val = f'{p_val}&"@@_SFX_99_@@"'
                                    c_bp = c_bp.replace(f"{{Param_{idx}}}", p_val)
                                    
                                for header_name, col_let in c_dz_col_map.items():
                                    c_bp = c_bp.replace(f"{{{header_name}}}", f"${col_let}{new_row_num}")
                                    
                                c_bp = apply_context_replacements(c_bp, c_ctx)
                                c_bp = c_bp.replace("@@_SFX_99_@@", "YTD")
                                return c_bp

                            shifted_formula = re.sub(DBRW_SUB_RGX, clone_universal_replacer, shifted_formula, flags=re.IGNORECASE)

                        c_f_tag.text = shifted_formula.lstrip('=')
                        
                        c_v_tag = c_cell.find('s:v', NS)
                        if c_v_tag is not None:
                            c_cell.remove(c_v_tag)
                        c_cell.attrib.pop('t', None)
                    else:
                        c_cell.remove(c_f_tag)
                        c_v_tag = c_cell.find('s:v', NS)
                        if c_v_tag is not None:
                            c_cell.remove(c_v_tag)
                        c_cell.attrib.pop('t', None)

        cells_c = [c for c in header_data_row if c.tag == f'{{{NS_URL}}}c']
        cells_c.sort(key=lambda c: col_letter_to_idx(parse_col_letter(c.attrib.get('r', 'A'))))
        for c in cells_c: header_data_row.remove(c)
        header_data_row.extend(cells_c)

        sheet_tree.append(header_data_row)

        for clone in cloned_rows_buffer:
            sheet_tree.append(clone)

        for r_elem in sheet_tree.findall('.//s:row', NS):
            for c_elem in list(r_elem):
                if c_elem.tag == f'{{{NS_URL}}}c' and len(list(c_elem)) == 0 and not c_elem.text and 's' not in c_elem.attrib:
                    r_elem.remove(c_elem)

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

    print("\n  ├─ Serializing sheets back to ZIP...")
    
    for sheet_path, tab_name in sheet_info:
        if tab_name in modified_tabs and tab_name in ram_sheets:
            header_text, footer_text = sheet_wrappers[tab_name]
            sheet_tree = ram_sheets[tab_name]
            modified_inner_str = ""
            for child in sheet_tree:
                modified_inner_str += ET.tostring(child, encoding='unicode')
            modified_inner_str = modified_inner_str.replace('ns0:', '').replace(':ns0', '')
            final_xml_str = header_text + modified_inner_str + footer_text
            output_zip.writestr(sheet_path, final_xml_str.encode('utf-8'))
        else:
            if tab_name in raw_bytes_map:
                output_zip.writestr(sheet_path, raw_bytes_map[tab_name])

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
            r_ids = [int(m) for m in re.findall(r'rId(\d+)"', rels_data)]
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
                
                col_let = get_col_letter(1)
                safe_tab_name = escape_xml(tab_n)
                audit_xml += f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{safe_tab_name}</t></is></c>\n'
                
                col_let = get_col_letter(2)
                audit_xml += f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{escape_xml(stats.get("Acct_Col", "N/A"))}</t></is></c>\n'
                
                col_let = get_col_letter(3)
                audit_xml += f'<c r="{col_let}{row_num}" t="inlineStr"><is><t>{escape_xml(stats.get("DZ_Cols", "N/A"))}</t></is></c>\n'
                
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

    try:
        yaml_output = "Cube_Cache_State:\n"
        for k, v in cube_cache.items(): yaml_output += f"  '{k}': '{v}'\n"
        
        yaml_output += "\nCube_Cache_Hits:\n"
        for k, v in cube_cache_hits.items(): yaml_output += f"  '{k}': {v}\n"
        
        yaml_output += "\nHeader_Cache_State:\n"
        for k, v in header_cache.items(): yaml_output += f"  '{k}': '{v}'\n"
        
        yaml_output += "\nHeader_Cache_Hits:\n"
        for k, v in header_cache_hits.items(): yaml_output += f"  '{k}': {v}\n"

        with open("Cache_Diagnostic_Log.yaml", "w", encoding="utf-8") as f:
            f.write(yaml_output)
        print("\n  ├─ [DEBUG] Exported Cache_Diagnostic_Log.yaml")
    except Exception as e:
        print(f"\n  └─ [DEBUG] Failed to write YAML log: {e}")

    output_zip.close()
    return output_buffer.getvalue()