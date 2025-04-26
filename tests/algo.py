import os
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from io import BytesIO

def extract_xlsx(xlsx_path):
    content = {}
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        for name in z.namelist():
            content[name] = z.read(name)
    return content

def parse_workbook(workbook_xml):
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    root = ET.fromstring(workbook_xml)
    sheets = []
    for sheet in root.findall('main:sheets/main:sheet', ns):
        sheets.append({
            'name': sheet.attrib['name'],
            'sheetId': int(sheet.attrib['sheetId']),
            'r:id': sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        })
    return sheets

def parse_relationships(rels_xml):
    root = ET.fromstring(rels_xml)
    relations = {}
    for rel in root.findall('Relationship', namespaces={'': "http://schemas.openxmlformats.org/package/2006/relationships"}):
        relations[rel.attrib['Id']] = {
            'Type': rel.attrib['Type'],
            'Target': rel.attrib['Target']
        }
    return relations

def resolve_sheet_names_conflict(sheet_names):
    counter = defaultdict(int)
    resolved = []
    for name in sheet_names:
        counter[name] += 1
        if counter[name] == 1:
            resolved.append(name)
        else:
            resolved.append(f"{name} ({counter[name]})")
    return resolved

def generate_new_rid(existing_rids):
    max_id = 0
    for rid in existing_rids:
        try:
            num = int(rid.replace('rId', ''))
            if num > max_id:
                max_id = num
        except:
            continue
    return f"rId{max_id + 1}"

def merge_workbooks(a_content, b_content):
    merged = {}
    merged.update(a_content)

    a_workbook = parse_workbook(a_content['xl/workbook.xml'])
    b_workbook = parse_workbook(b_content['xl/workbook.xml'])

    a_rels = parse_relationships(a_content['xl/_rels/workbook.xml.rels'])
    b_rels = parse_relationships(b_content['xl/_rels/workbook.xml.rels'])

    merged_rels = dict(a_rels)
    merged_sheets = list(a_workbook)
    existing_rids = set(merged_rels.keys())
    existing_sheet_ids = {s['sheetId'] for s in merged_sheets}

    for sheet in b_workbook:
        old_rid = sheet['r:id']
        new_rid = generate_new_rid(existing_rids)
        existing_rids.add(new_rid)

        # Copiar sheet XML
        old_target = b_rels[old_rid]['Target']
        if old_target.startswith('/'):
            old_target = old_target[1:]
        sheet_xml = b_content[f'xl/{old_target}']
        sheet_target = f'xl/worksheets/sheet{sheet["sheetId"]+1000}.xml'  # Evitar conflicto de nombres
        merged[sheet_target] = sheet_xml

        # Crear nueva relación
        merged_rels[new_rid] = {
            'Type': "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            'Target': f'worksheets/sheet{sheet["sheetId"]+1000}.xml'
        }

        # Agregar nuevo sheet
        merged_sheets.append({
            'name': sheet['name'],
            'sheetId': max(existing_sheet_ids)+1,
            'r:id': new_rid
        })
        existing_sheet_ids.add(max(existing_sheet_ids)+1)

    final_sheet_names = resolve_sheet_names_conflict([s['name'] for s in merged_sheets])
    for idx, name in enumerate(final_sheet_names):
        merged_sheets[idx]['name'] = name

    merged['xl/workbook.xml'] = build_new_workbook(merged_sheets)
    merged['xl/_rels/workbook.xml.rels'] = build_new_workbook_rels(merged_rels)
    merged['docProps/app.xml'] = build_new_app_xml(final_sheet_names)
    merged['docProps/core.xml'] = build_new_core_xml('Merged by Script')

    return merged

def build_new_workbook(sheet_definitions):
    nsmap = {
        None: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    }
    ET.register_namespace('', nsmap[None])
    ET.register_namespace('r', nsmap['r'])

    workbook = ET.Element('workbook')
    sheets = ET.SubElement(workbook, 'sheets')

    for sheet in sheet_definitions:
        sheet_elem = ET.SubElement(sheets, 'sheet', {
            'name': sheet['name'],
            'sheetId': str(sheet['sheetId']),
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': sheet['r:id']
        })

    tree = ET.ElementTree(workbook)
    output = BytesIO()
    tree.write(output, encoding='utf-8', xml_declaration=True)
    return output.getvalue()

def build_new_workbook_rels(all_relationships):
    ET.register_namespace('', "http://schemas.openxmlformats.org/package/2006/relationships")
    relationships = ET.Element('Relationships', xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
    for rid, rel in all_relationships.items():
        ET.SubElement(relationships, 'Relationship', {
            'Id': rid,
            'Type': rel['Type'],
            'Target': rel['Target']
        })
    tree = ET.ElementTree(relationships)
    output = BytesIO()
    tree.write(output, encoding='utf-8', xml_declaration=True)
    return output.getvalue()

def build_new_app_xml(sheet_names):
    ET.register_namespace('', "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
    ET.register_namespace('vt', "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
    props = ET.Element('Properties', {
        'xmlns': "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        'xmlns:vt': "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    })
    ET.SubElement(props, 'Application').text = 'Microsoft Excel'
    heading_pairs = ET.SubElement(props, 'HeadingPairs')
    vec = ET.SubElement(heading_pairs, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}vector', {
        'size': '2', 'baseType': 'variant'
    })
    var1 = ET.SubElement(vec, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}variant')
    ET.SubElement(var1, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpstr').text = 'Worksheets'
    var2 = ET.SubElement(vec, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}variant')
    ET.SubElement(var2, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}i4').text = str(len(sheet_names))

    titles = ET.SubElement(props, 'TitlesOfParts')
    t_vec = ET.SubElement(titles, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}vector', {
        'size': str(len(sheet_names)), 'baseType': 'lpstr'
    })
    for name in sheet_names:
        ET.SubElement(t_vec, '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpstr').text = name

    tree = ET.ElementTree(props)
    output = BytesIO()
    tree.write(output, encoding='utf-8', xml_declaration=True)
    return output.getvalue()

def build_new_core_xml(creator):
    ET.register_namespace('cp', "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
    ET.register_namespace('dc', "http://purl.org/dc/elements/1.1/")
    ET.register_namespace('dcterms', "http://purl.org/dc/terms/")
    ET.register_namespace('dcmitype', "http://purl.org/dc/dcmitype/")
    ET.register_namespace('xsi', "http://www.w3.org/2001/XMLSchema-instance")
    props = ET.Element('cp:coreProperties')
    ET.SubElement(props, 'dc:creator').text = creator
    ET.SubElement(props, 'cp:lastModifiedBy').text = creator
    ET.SubElement(props, 'dcterms:created', {'xsi:type': 'dcterms:W3CDTF'}).text = '2025-04-25T00:00:00Z'
    ET.SubElement(props, 'dcterms:modified', {'xsi:type': 'dcterms:W3CDTF'}).text = '2025-04-25T00:00:00Z'
    tree = ET.ElementTree(props)
    output = BytesIO()
    tree.write(output, encoding='utf-8', xml_declaration=True)
    return output.getvalue()

def write_xlsx(output_path, merged_content):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for path, data in merged_content.items():
            z.writestr(path, data)

if __name__ == "__main__":
    a_file = "Example A.xlsx"
    b_file = "Example C.xlsx"
    output_file = "Merged_Output.xlsx"

    a_content = extract_xlsx(a_file)
    b_content = extract_xlsx(b_file)
    merged_content = merge_workbooks(a_content, b_content)
    write_xlsx(output_file, merged_content)

    print(f"✅ Archivo fusionado creado: {output_file}")
