import shutil
import os
from core.xml_utils import parse_xml, write_xml, generate_unique_id

class SheetManager:
    def __init__(self, sheets_dir, workbook_path):
        self.sheets_dir = sheets_dir
        self.workbook_path = workbook_path

    def list_sheets(self):
        _, root = parse_xml(self.workbook_path)
        sheets = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets')
        return [sheet.attrib['name'] for sheet in sheets.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')]

    def add_sheet_entry(self, sheet_id, r_id, name):
        tree, root = parse_xml(self.workbook_path)
        sheets = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets')
        sheet = ET.Element('sheet', name=name, sheetId=str(sheet_id), attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': r_id})
        sheets.append(sheet)
        write_xml(tree, self.workbook_path)

    def copy_sheet_file(self, src_sheet_path, dst_sheet_path):
        shutil.copy(src_sheet_path, dst_sheet_path)
