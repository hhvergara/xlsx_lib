from core.zip_handler import ZipHandler
from core.xml_utils import parse_xml_bytes, write_xml_to_bytes, generate_unique_id
import xml.etree.ElementTree as ET

class XLSXMerger:
    def __init__(self, file_a_bytes, file_b_bytes):
        self.zip_a = ZipHandler(file_a_bytes)
        self.zip_b = ZipHandler(file_b_bytes)
        self.output_zip = ZipHandler()

    def merge(self):
        self.zip_a.extract()
        self.zip_b.extract()

        self.output_zip.files = dict(self.zip_a.files)

        self.merge_sheets()
        # self.ensure_docProps()
        self.update_app_xml()
        self.update_core_xml()
        self.ensure_content_types_and_rels()

        return self.output_zip.create_zip_bytes()

    def merge_sheets(self):
        output_workbook = self.output_zip.get_file_content('xl/workbook.xml')
        output_workbook_rels = self.output_zip.get_file_content('xl/_rels/workbook.xml.rels')

        src_workbook = self.zip_b.get_file_content('xl/workbook.xml')
        src_workbook_rels = self.zip_b.get_file_content('xl/_rels/workbook.xml.rels')

        tree_out, root_out = parse_xml_bytes(output_workbook)
        tree_src, root_src = parse_xml_bytes(src_workbook)

        sheets_out = root_out.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets')
        sheets_src = root_src.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets')

        existing_names = {sheet.attrib['name'] for sheet in sheets_out.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')}
        sheet_id = len(existing_names) + 1

        for sheet in sheets_src.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
            original_name = sheet.attrib['name']
            new_name = original_name
            while new_name in existing_names:
                new_name = f"{original_name}(1)"
                original_name += "(1)"
            existing_names.add(new_name)

            new_sheet_filename = f"xl/worksheets/sheet{sheet_id}.xml"
            old_rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']

            # Copy the sheet XML
            old_sheet_target = self.find_target_in_rels(self.zip_b.get_file_content('xl/_rels/workbook.xml.rels'), old_rid)
            sheet_bytes = self.zip_b.get_file_content(f"xl/{old_sheet_target}")
            self.output_zip.set_file_content(new_sheet_filename, sheet_bytes)

            # Create new rel ID
            new_rid = generate_unique_id(self.get_existing_rids(output_workbook_rels))
            self.append_to_rels(output_workbook_rels, new_rid, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', f"worksheets/sheet{sheet_id}.xml")

            # Append sheet
            sheet.attrib['name'] = new_name
            sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'] = new_rid
            sheets_out.append(sheet)
            sheet_id += 1

        self.output_zip.set_file_content('xl/workbook.xml', write_xml_to_bytes(tree_out))
        self.output_zip.set_file_content('xl/_rels/workbook.xml.rels', output_workbook_rels)

    def ensure_docProps(self):
        if "docProps/app.xml" not in self.output_zip.files:
            self.output_zip.set_file_content("docProps/app.xml", b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>Python XLSXMerger</Application>
</Properties>''')
        if "docProps/core.xml" not in self.output_zip.files:
            self.output_zip.set_file_content("docProps/core.xml", b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>Python XLSXMerger</dc:creator>
    <cp:lastModifiedBy>Python XLSXMerger</cp:lastModifiedBy>
</cp:coreProperties>''')

    def ensure_content_types_and_rels(self):
        if "[Content_Types].xml" not in self.output_zip.files:
            self.output_zip.set_file_content("[Content_Types].xml", b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
</Types>''')
        if "_rels/.rels" not in self.output_zip.files:
            self.output_zip.set_file_content("_rels/.rels", b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>''')

    def find_target_in_rels(self, rels_bytes, r_id):
        tree, root = parse_xml_bytes(rels_bytes)
        for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            if rel.attrib['Id'] == r_id:
                return rel.attrib['Target']
        return None

    def get_existing_rids(self, rels_bytes):
        tree, root = parse_xml_bytes(rels_bytes)
        return {rel.attrib['Id'] for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')}

    def append_to_rels(self, rels_bytes, new_id, type_, target):
        tree, root = parse_xml_bytes(rels_bytes)
        ET.SubElement(root, 'Relationship', Id=new_id, Type=type_, Target=target)
        self.output_zip.set_file_content('xl/_rels/workbook.xml.rels', write_xml_to_bytes(tree))


    def update_app_xml(self):
        sheet_names = [x for x in self.output_zip.files if 'xl/worksheets/' in x]
        root = ET.Element('Properties', {
            'xmlns': "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
            'xmlns:vt': "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        })
        ET.SubElement(root, 'Template').text = ''
        ET.SubElement(root, 'Application').text = 'Microsoft Excel'
        ET.SubElement(root, 'DocSecurity').text = '0'
        ET.SubElement(root, 'ScaleCrop').text = 'false'

        heading_pairs = ET.SubElement(root, 'HeadingPairs')
        hp_vec = ET.SubElement(heading_pairs, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}vector',
                                {'size': '2', 'baseType': 'variant'})
        var1 = ET.SubElement(hp_vec, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}variant')
        ET.SubElement(var1, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpstr').text = 'Hojas de cálculo'
        var2 = ET.SubElement(hp_vec, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}variant')
        ET.SubElement(var2, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}i4').text = str(len(sheet_names))

        titles = ET.SubElement(root, 'TitlesOfParts')
        t_vec = ET.SubElement(titles, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}vector',
                                {'size': str(len(sheet_names)), 'baseType': 'lpstr'})
        for name in sheet_names:
            ET.SubElement(t_vec, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpstr').text = name

        ET.SubElement(root, 'Manager').text = 'MGE'
        ET.SubElement(root, 'Company').text = 'CVS'
        ET.SubElement(root, 'LinksUpToDate').text = 'false'
        ET.SubElement(root, 'SharedDoc').text = 'false'
        ET.SubElement(root, 'HyperlinksChanged').text = 'false'
        ET.SubElement(root, 'AppVersion').text = '16.0300'

        self.output_zip.set_file_content('docProps/app.xml', b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root, encoding='utf-8', method='xml'))

    def update_core_xml(self, creator='mge'):
        root = ET.Element('cp:coreProperties', {
            'xmlns:cp': "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            'xmlns:dc': "http://purl.org/dc/elements/1.1/",
            'xmlns:dcterms': "http://purl.org/dc/terms/",
            'xmlns:dcmitype': "http://purl.org/dc/dcmitype/",
            'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance"
        })
        for tag in ['title', 'subject', 'keywords', 'description', 'category']:
            ET.SubElement(root, f'dc:{tag}').text = tag
        ET.SubElement(root, 'dc:creator').text = creator
        ET.SubElement(root, 'cp:lastModifiedBy').text = creator
        ET.SubElement(root, 'dcterms:created', {'xsi:type': "dcterms:W3CDTF"}).text = '2025-04-25T18:00:00Z'
        ET.SubElement(root, 'dcterms:modified', {'xsi:type': "dcterms:W3CDTF"}).text = '2025-04-25T18:00:00Z'

        self.output_zip.set_file_content('docProps/core.xml', b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root, encoding='utf-8', method='xml'))
