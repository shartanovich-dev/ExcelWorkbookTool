import re
import xml.etree.ElementTree as ET

# XML Namespaces
NS_URL = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS = {'s': NS_URL}
ET.register_namespace('', NS_URL)


# --- CELL ADDRESS HELPERS ---
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
    if not cell_ref:
        return "A"
    clean = str(cell_ref).split('!')[-1].replace('$', '')
    match = re.match(r"([A-Z]+)", clean, re.IGNORECASE)
    return match.group(1).upper() if match else "A"

def parse_row_num(cell_ref):
    if not cell_ref:
        return "1"
    clean = str(cell_ref).split('!')[-1].replace('$', '')
    match = re.match(r"[A-Z]+(\d+)", clean, re.IGNORECASE)
    return match.group(1) if match else "1"


# --- WORKBOOK READING HELPERS ---
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
