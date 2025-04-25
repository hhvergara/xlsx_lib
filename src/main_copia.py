import os
import zipfile
from io import BytesIO
from pathlib import Path
import xml.etree.ElementTree as ET

__BASE_DIR__ = Path(__file__).resolve().parent

def extract_xml_from_xlsx(xlsx_path):
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        xml_files = {name: z.read(name) for name in z.namelist()}
    return xml_files

def get_relations(xml_data):
    relations = {}
    root = [x for x in ET.fromstring(xml_data)]
    for rel in root:#.findall('.//Relationships'):
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

def write_xlsx(output_path, xml_data, relations):
    with zipfile.ZipFile(output_path, 'w') as z:
        for filename, xml in xml_data.items():
            z.writestr(filename, xml)
        # Add the updated relations
        relationships_xml = generate_relationships_xml(relations)
        z.writestr('xl/_rels/.rels', relationships_xml)

def generate_relationships_xml(relations):
    # Generate new rels.xml
    root = ET.Element('Relationships', xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
    for rid, target in relations.items():
        rel_elem = ET.SubElement(root, 'Relationship')
        rel_elem.set('Id', rid)
        rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet')
        rel_elem.set('Target', target)
    xml_bytes = ET.tostring(root, encoding='utf-8', method='xml')
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_bytes


def merge_xlsx_files(a_path, b_path, output_path):
    # Extract XML from both files
    a_xml = extract_xml_from_xlsx(a_path)
    b_xml = extract_xml_from_xlsx(b_path)

    # Initialize the list for new rIds
    new_rids = {}

    # Process A.xlsx first
    a_relations = get_relations(a_xml.get('xl/_rels/workbook.xml.rels'))
    for rid in a_relations:
        new_rid = create_new_rid(new_rids)
        new_rids[rid] = new_rid

    # Update A's XML content with new rIds
    a_xml['xl/workbook.xml'] = update_relations(a_xml['xl/workbook.xml'], new_rids)

    # Create Output.xlsx from A's XML data
    write_xlsx(output_path, a_xml, new_rids)

    # Process B.xlsx, continuing the rId sequence
    b_relations = get_relations(b_xml.get('xl/_rels/workbook.xml.rels'))
    for rid in b_relations:
        new_rid = create_new_rid(new_rids)
        new_rids[rid] = new_rid

    # Add B's XML to Output.xlsx and update relations
    b_xml['xl/workbook.xml'] = update_relations(b_xml['xl/workbook.xml'], new_rids)
    write_xlsx(output_path, b_xml, new_rids)

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
