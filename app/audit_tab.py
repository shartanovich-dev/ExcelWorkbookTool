"""Adds the 'Conversion Audit' sheet to a converted workbook."""
import re

from app.xlsx_helpers import NS_URL, get_col_letter


def add_audit_tab(output_zip, structural_files, audit_log):
    """Write the workbook's structural files into output_zip, adding a 'Conversion Audit'
    sheet built from audit_log. If anything fails, the original structural files are written."""
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

            headers = ['Sheet Name', 'Account Column', 'Dropzone Column(s)', 'Total Evaluated', 'Converted (Success)', 'Exceptions', 'Incomplete Match', 'Unmapped (Orphans)', 'Duplicates Suppressed', 'Unlocked Refs']
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
                    stats['Duplicates'],
                    stats.get('Unlocked_Refs', 0)
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
