import os
import zipfile
from io import BytesIO
from pathlib import Path
import xml.etree.ElementTree as ET
from collections import defaultdict

__BASE_DIR__ = Path(__file__).resolve().parent

def extract_xml_from_xlsx(xlsx_path):
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        xml_files = {name: z.read(name) for name in z.namelist()}
    return xml_files

def get_sheet_names(workbook_xml):
    ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    root = ET.fromstring(workbook_xml)
    return [sheet.attrib['name'] for sheet in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')]

def get_relations(xml_data):
    relations = {}
    root = [x for x in ET.fromstring(xml_data)]
    for rel in root:
        rId = rel.attrib['Id']
        target = rel.attrib['Target']
        relations[rId] = target
    return relations

def update_relations(xml_data, new_rids):
    root = ET.fromstring(xml_data)
    for rel in root.findall('.//Relationship'):
        rId = rel.attrib['Id']
        if rId in new_rids:
            rel.attrib['Id'] = new_rids[rId]
    xml_bytes = ET.tostring(root, encoding='utf-8', method='xml')
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_bytes

def create_new_rid(new_rids):
    return f"rId{len(new_rids) + 1}"

def update_app_xml(sheet_names):
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

    ET.SubElement(root, 'Manager').text = ''
    ET.SubElement(root, 'Company').text = ''
    ET.SubElement(root, 'LinksUpToDate').text = 'false'
    ET.SubElement(root, 'SharedDoc').text = 'false'
    ET.SubElement(root, 'HyperlinksChanged').text = 'false'
    ET.SubElement(root, 'AppVersion').text = '16.0300'

    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root, encoding='utf-8', method='xml')

def update_core_xml(creator=''):
    root = ET.Element('cp:coreProperties', {
        'xmlns:cp': "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        'xmlns:dc': "http://purl.org/dc/elements/1.1/",
        'xmlns:dcterms': "http://purl.org/dc/terms/",
        'xmlns:dcmitype': "http://purl.org/dc/dcmitype/",
        'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance"
    })
    for tag in ['title', 'subject', 'keywords', 'description', 'category']:
        ET.SubElement(root, f'dc:{tag}').text = ''
    ET.SubElement(root, 'dc:creator').text = creator
    ET.SubElement(root, 'cp:lastModifiedBy').text = creator
    ET.SubElement(root, 'dcterms:created', {'xsi:type': "dcterms:W3CDTF"}).text = '2025-04-25T18:00:00Z'
    ET.SubElement(root, 'dcterms:modified', {'xsi:type': "dcterms:W3CDTF"}).text = '2025-04-25T18:00:00Z'

    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root, encoding='utf-8', method='xml')

def write_xlsx(output_path, xml_data, relations, sheet_names, creator=''):
    with zipfile.ZipFile(output_path, 'w') as z:
        for filename, xml in xml_data.items():
            z.writestr(filename, xml)
        relationships_xml = generate_relationships_xml(relations)
        z.writestr('xl/_rels/.rels', relationships_xml)
        z.writestr('docProps/app.xml', update_app_xml(sheet_names))
        z.writestr('docProps/core.xml', update_core_xml(creator=creator))

def generate_relationships_xml(relations):
    root = ET.Element('Relationships', xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
    for rid, target in relations.items():
        rel_elem = ET.SubElement(root, 'Relationship')
        rel_elem.set('Id', rid)
        rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet')
        rel_elem.set('Target', target)
    xml_bytes = ET.tostring(root, encoding='utf-8', method='xml')
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_bytes

def merge_xlsx_files(a_path, b_path, output_path):
    a_xml = extract_xml_from_xlsx(a_path)
    b_xml = extract_xml_from_xlsx(b_path)
    new_rids = {}

    a_relations = get_relations(a_xml.get('xl/_rels/workbook.xml.rels'))
    for rid in a_relations:
        new_rid = create_new_rid(new_rids)
        new_rids[rid] = new_rid
    a_xml['xl/workbook.xml'] = update_relations(a_xml['xl/workbook.xml'], new_rids)

    sheet_names_a = get_sheet_names(a_xml['xl/workbook.xml'])
    name_counter = defaultdict(int)
    final_names = []

    for name in sheet_names_a:
        name_counter[name] += 1
        final_names.append(name if name_counter[name] == 1 else f"{name} ({name_counter[name]})")

    b_relations = get_relations(b_xml.get('xl/_rels/workbook.xml.rels'))
    for rid in b_relations:
        new_rid = create_new_rid(new_rids)
        new_rids[rid] = new_rid
    b_xml['xl/workbook.xml'] = update_relations(b_xml['xl/workbook.xml'], new_rids)

    sheet_names_b = get_sheet_names(b_xml['xl/workbook.xml'])
    for name in sheet_names_b:
        name_counter[name] += 1
        final_names.append(name if name_counter[name] == 1 else f"{name} ({name_counter[name]})")

    all_xml = {**a_xml, **b_xml}
    write_xlsx(output_path, all_xml, new_rids, final_names, creator='')

# Descomprimir y recomprimir siguen igual, no los incluyo por espacio
def descomprimir_xlsx(ruta_xlsx, carpeta_destino):
    Path(carpeta_destino).mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(ruta_xlsx, 'r') as zipf:
        zipf.extractall(carpeta_destino)
    print(f"✅ Extraído en: {carpeta_destino}")


def recomprimir_a_xlsx(carpeta_origen, ruta_salida_xlsx):
    with zipfile.ZipFile(ruta_salida_xlsx, "w", zipfile.ZIP_DEFLATED) as zipf:
        for foldername, _, filenames in os.walk(carpeta_origen):
            for filename in filenames:
                file_path = Path(foldername) / filename
                arcname = file_path.relative_to(carpeta_origen)
                zipf.write(file_path, arcname)
    print(f"✅ Reempaquetado como: {ruta_salida_xlsx}")


# Ejemplo de uso
if __name__ == '__main__':
    file_1 = os.path.join(__BASE_DIR__,'examples', 'Example A.xlsx')
    file_2 = os.path.join(__BASE_DIR__,'examples', 'Example C.xlsx')
    output_filename = "Output.xlsx"
    output_folder_file = os.path.join(__BASE_DIR__,'output', output_filename)

    merge_xlsx_files(file_1, file_2, output_folder_file)

    descomprimir_xlsx(
            ruta_xlsx=output_folder_file,
            carpeta_destino=output_folder_file.replace('.xlsx', '')
        )
    import shutil
    shutil.copy2(output_folder_file, output_folder_file.replace('.xlsx', '.zip'))


    recomprimir_a_xlsx(carpeta_origen=output_folder_file.replace('.xlsx', 'Output'), ruta_salida_xlsx=output_folder_file.replace('.xlsx', '_recomprimido.xlsx'))
