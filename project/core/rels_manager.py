import os
import xml.etree.ElementTree as ET
from core.xml_utils import parse_xml, write_xml, generate_unique_id

class RelsManager:
    def __init__(self, rels_path):
        self.rels_path = rels_path
        self.tree, self.root = parse_xml(rels_path)

    def list_relationships(self):
        return [rel.attrib for rel in self.root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')]

    def add_relationship(self, Id, Type, Target):
        rel = ET.Element('Relationship', Id=Id, Type=Type, Target=Target)
        self.root.append(rel)

    def save(self):
        write_xml(self.tree, self.rels_path)

    def get_existing_ids(self):
        return {rel.attrib['Id'] for rel in self.root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')}

    def generate_new_id(self):
        return generate_unique_id(self.get_existing_ids())
