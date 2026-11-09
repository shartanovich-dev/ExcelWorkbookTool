import re

from app.xlsx_helpers import NS

# --- BULLETPROOF EXCEL REGEX PATTERNS (DYNAMIC TO PREVENT UI CORRUPTION) ---
PIPE = chr(124)
DBRW_EXTRACT_RGX = r'(?:_xll\.)?DBRW\s*\(((?:[^)("\']' + PIPE + r'"[^"]*"' + PIPE + r'\'[^\']*\'' + PIPE + r'\([^)]*\))*)\)'
DBRW_SUB_RGX = r'@?(?:_xll\.)?DBRW\s*\((?:[^)("\']' + PIPE + r'"[^"]*"' + PIPE + r'\'[^\']*\'' + PIPE + r'\([^)]*\))*\)'


# --- FORMULA HELPERS ---
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
