import xml.etree.ElementTree as ET
import io

def parse_xml_bytes(xml_bytes):
    tree = ET.ElementTree(ET.fromstring(xml_bytes))
    return tree, tree.getroot()

def write_xml_to_bytes(tree):
    output = io.BytesIO()
    tree.write(output, encoding="utf-8", xml_declaration=True)
    return output.getvalue()

def generate_unique_id(existing_ids, prefix='rId'):
    """
    Genera un identificador único que no esté en existing_ids.
    """
    counter = 1
    while f"{prefix}{counter}" in existing_ids:
        counter += 1
    return f"{prefix}{counter}"