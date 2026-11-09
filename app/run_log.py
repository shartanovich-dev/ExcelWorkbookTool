"""Diagnostic logging for a conversion run."""
from app.config import DIAG_LOG_PATH


def build_diagnostic_yaml(master_category_map, cube_cache, cube_cache_hits, header_cache, header_cache_hits):
    """Return the cache diagnostics of one workbook conversion as YAML text."""
    yaml_output = "Category_Map_State:\n"
    for k, v in master_category_map.items(): yaml_output += f"  '{k}': '{v}'\n"

    yaml_output += "Cube_Cache_State:\n"
    for k, v in cube_cache.items(): yaml_output += f"  '{k}': '{v}'\n"

    yaml_output += "\nCube_Cache_Hits:\n"
    for k, v in cube_cache_hits.items(): yaml_output += f"  '{k}': {v}\n"

    yaml_output += "\nHeader_Cache_State:\n"
    for k, v in header_cache.items(): yaml_output += f"  '{k}': '{v}'\n"

    yaml_output += "\nHeader_Cache_Hits:\n"
    for k, v in header_cache_hits.items(): yaml_output += f"  '{k}': {v}\n"
    return yaml_output


def write_diagnostic_log(yaml_text):
    try:
        with open(DIAG_LOG_PATH, "w", encoding="utf-8") as f:
            f.write(yaml_text)
        print(f"\n  ├─ [DEBUG] Exported {DIAG_LOG_PATH}")
    except Exception as e:
        print(f"\n  └─ [DEBUG] Failed to write YAML log: {e}")
