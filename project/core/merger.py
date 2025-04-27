from typing import Optional
import zipfile
from core.zip_handler import ZipHandler
from core.xml_utils import parse_xml_bytes, write_xml_to_bytes, generate_unique_id
import xml.etree.ElementTree as ET

class XLSXMerger:
    def __init__(self, file_a_bytes, file_b_bytes):
        self.zip_a = ZipHandler(file_a_bytes)
        self.zip_b = ZipHandler(file_b_bytes)
        self.output_zip = ZipHandler()


    def merge(self):
        """
        Merge the two XLSX files.
        """
        # Extract the contents of both ZIP files
        self.zip_a.extract()
        self.zip_b.extract()

        # Genero un archivo output y copio A
        self.output_zip.files = self.zip_a.files.copy()
        self.update_filenames()

        # obtengo las relaciones de ambos archivos
        relations_a = self.get_relations(self.zip_a)
        relations_b = self.get_relations(self.zip_b)
        relations_output = self.get_relations(self.output_zip)

        # # Me fijo que relaciones se repiten en B:
        # pepe = self.list_files()
        # to_update = []
        # for rel in relations_b:
        #     if rel in relations_output and rel not in '_rels/.rels':
        #         to_update.append(rel)
        # # Actualizo las relaciones de B en el output
        # for rel in to_update:
        #     b_relations_to_check = relations_b[rel]['relations']
        #     # Obtengo el maximo id de las relaciones de Output
        #     max_a_id = int(max(relations_a[rel]['relations'].keys()).replace('rId', ''))

        #     a_relations_to_check = relations_b[rel]['relations']
        #     relation_mapping = {}
        #     for rid, target in b_relations_to_check.items():
        #         # Genero un nuevo id para la relacion
        #         new_rid = f"rId{max_a_id + 1}"
        #         relation_mapping[rid] = new_rid
        #         max_a_id += 1
        #     # obtengo el contenido a actualizar
        #     content = self.zip_b.get_file_content(rel)
        #     tree, root = parse_xml_bytes(content)
        #     # Actualizo los ids de las relaciones
        #     for _rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        #         rid = _rel.get('Id')
        #         if rid in relation_mapping:
        #             _rel.set('Id', relation_mapping[rid])
        #             self.zip_b.set_file_content(rel, write_xml_to_bytes(tree))
        #     relations_b = self.get_relations(self.zip_b)


        # a_app_xml = self.get_app_xml(self.zip_a)
        # b_app_xml = self.get_app_xml(self.zip_b)
        # a_core_xml = self.get_core_xml(self.zip_a)
        # b_core_xml = self.get_core_xml(self.zip_b)
        # a_styles_xml = self.get_styles(self.zip_a)
        # b_styles_xml = self.get_styles(self.zip_b)
        # a_shared_strings_xml = self.get_shared_strings(self.zip_a)
        # b_shared_strings_xml = self.get_shared_strings(self.zip_b)
        # a_worksheets_xml = self.get_worksheets(self.zip_a)
        # b_worksheets_xml = self.get_worksheets(self.zip_b)


        # Merge the contents
        # Write the merged contents to the output ZIP file
        with zipfile.ZipFile('output_path.xlsx', 'w') as z:
            out_files = self.output_zip.list_files()
            b_files = self.zip_b.list_files()
            merged_files = ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels']
            self.output_zip.files['xl/workbook.xml'] = self.merge_all_nested_deep(self.output_zip.get_file_content('xl/workbook.xml'), self.zip_b.get_file_content('xl/workbook.xml'))
            self.output_zip.files['xl/_rels/workbook.xml.rels'] = self.merge_all_nested_deep(self.output_zip.get_file_content('xl/_rels/workbook.xml.rels'), self.zip_b.get_file_content('xl/workbook.xml'))
            
            self.output_zip.files['xl/workbook.xml'] = self.merge_workbook_strings(
                self.output_zip.get_file_content('xl/workbook.xml'),
                self.zip_b.get_file_content('xl/workbook.xml')
            )
            self.output_zip.files['xl/_rels/workbook.xml.rels'] = self.merge_relationships_strings(
                self.output_zip.get_file_content('xl/_rels/workbook.xml.rels'),
                self.zip_b.get_file_content('xl/_rels/workbook.xml.rels')
            )
            
            

            self.listar_sheets(self.output_zip.get_file_content('xl/workbook.xml'))
 
            self.listar_relationships(self.output_zip.get_file_content('xl/_rels/workbook.xml.rels'))
            for filename in out_files:
                z.writestr(filename, self.output_zip.get_file_content(filename))
                merged_files.append(filename)
            for filename in b_files:
                if filename not in merged_files:
                    z.writestr(filename, self.zip_b.get_file_content(filename))
                    merged_files.append(filename)

        # Create the output ZIP file
        merged_zip_bytes = self.output_zip.create_zip_bytes()
        return merged_zip_bytes

    def get_sheets(self, file):
        """
        Get the list of sheets from the workbook.
        """
        workbook_xml = file.get_file_content('xl/workbook.xml')
        tree, root = parse_xml_bytes(workbook_xml)
        sheets = []
        for sheet in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
            sheet_name = sheet.get('name')
            sheet_id = sheet.get('sheetId')
            sheets.append((sheet_name, sheet_id))
        return sheets

    def get_relations(self, file, _type: str = 'xl/_rels/workbook.xml.rels' ):
        """
        Get the list of relationships from the workbook.
        """
        rels = {}
        for rel in file.list_files():
            if rel.endswith('.rels'):

                workbook_rels_xml = file.get_file_content(rel)
                tree, root = parse_xml_bytes(workbook_rels_xml)

                relations = {}
                for _rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rid = _rel.get('Id')
                    target = _rel.get('Target')
                    relations[rid] = target
                rels.update({
                    rel:{
                    'content': workbook_rels_xml,
                    'tree': tree,
                    'root': root,
                    'relations': relations
                }})
        return rels


    def list_files(self, xml_path: Optional[str] = None) -> dict:
        """
        List all files in the ZIP.
        """
        if xml_path:
            return {'file_a': self.zip_a.list_files(xml_path), 'file_b': self.zip_b.list_files(xml_path)}
        else:
            return {'file_a': self.zip_a.list_files(), 'file_b': self.zip_b.list_files()}


    def update_filenames(self):
        files = self.list_files()
        charts_rels = []
        charts = []
        drawings_rels = []
        drawings = []
        media = []
        chart_colors = []
        chart_styles = []
        to_update = {}

        for element in files['file_a']:
            charts.append(element) if element.startswith('xl/charts/chart') and not 'rels' in element else None
            charts_rels.append(element) if element.startswith('xl/charts/chart/_rels/') else None
            chart_colors.append(element) if element.startswith('xl/charts/colors') else None
            chart_styles.append(element) if element.startswith('xl/charts/style') else None
            drawings_rels.append(element) if element.startswith('xl/drawings/_rels/') else None
            drawings.append(element) if element.startswith('xl/drawings/') and not 'rels' in element else None
            media.append(element) if element.startswith('xl/media/') else None

        max_charts_rels = max([int(x.split('xl/charts/_rels/chart')[-1].split('.xml.rels')[0]) for x in charts_rels]) if charts_rels else 0
        max_charts = max([int(x.split('xl/charts/chart')[-1].split('.xml')[0]) for x in charts]) if charts else 0
        max_chart_colors = max([int(x.split('xl/charts/colors')[-1].split('.xml')[0]) for x in chart_colors]) if chart_colors else 0
        max_chart_styles = max([int(x.split('xl/charts/style')[-1].split('.xml')[0]) for x in chart_styles]) if chart_styles else 0
        max_drawings_rels = max([int(x.split('xl/drawings/_rels/drawing')[-1].split('.xml.rels')[0]) for x in drawings_rels]) if drawings_rels else 0
        max_drawing = max([int(x.split('xl/drawings/drawing')[-1].split('.xml')[0]) for x in drawings]) if drawings else 0

        b_charts = []
        b_charts_rels = []
        b_chart_colors = []
        b_chart_styles = []
        b_drawings_rels = []
        b_drawings = []
        b_media = []

        for element in files['file_b']:
            b_charts.append(element) if element.startswith('xl/charts/chart') and not 'rels' in element else None
            b_charts_rels.append(element) if element.startswith('xl/charts/chart/_rels/') else None
            b_chart_colors.append(element) if element.startswith('xl/charts/colors') else None
            b_chart_styles.append(element) if element.startswith('xl/charts/style') else None
            b_drawings_rels.append(element) if element.startswith('xl/drawings/_rels/') else None
            b_drawings.append(element) if element.startswith('xl/drawings/') and not 'rels' in element else None
            b_media.append(element) if element.startswith('xl/media/') else None

        if any(x in charts for x in b_charts):
            for x in sorted(b_charts):
                to_update[x] = f"xl/charts/chart{(max_charts := max_charts + 1)}.xml"
        if any(x in charts_rels for x in b_charts_rels):
            for x in sorted(b_charts_rels):
                to_update[x] = f"xl/charts/_rels/chart{(max_charts_rels := max_charts_rels + 1)}.xml.rels"
        if any(x in drawings for x in b_drawings):
            for x in sorted(b_drawings):
                to_update[x] = f"xl/drawings/drawing{(max_drawing := max_drawing + 1)}.xml"
        if any(x in drawings_rels for x in b_drawings_rels):
            for x in sorted(b_drawings_rels):
                to_update[x] = f"xl/drawings/_rels/drawing{(max_drawings_rels := max_drawings_rels + 1)}.xml.rels"
        if any(x in chart_colors for x in b_chart_colors):
            for x in sorted(b_chart_colors):
                to_update[x] = f"xl/charts/colors{(max_chart_colors := max_chart_colors + 1)}.xml"
        if any(x in chart_styles for x in b_chart_styles):
            for x in sorted(b_chart_styles):
                to_update[x] = f"xl/charts/style{(max_chart_styles := max_chart_styles + 1)}.xml"
        if any(x in media for x in b_media):
            for x in sorted(b_media):
                to_update[x] = f"{x.split('.')[:-1][0]}(1).{x.split('.')[-1]}"

        for filename, new_filename in to_update.items():
            self.zip_b.files[new_filename] = self.zip_b.files.pop(filename)

        out_workbook_sheets = self.get_workbook_sheets(self.output_zip.get_file_content('xl/workbook.xml'))
        b_workbook_sheets = self.get_workbook_sheets(self.zip_b.get_file_content('xl/workbook.xml'))
        max_sheets =  len(out_workbook_sheets)
        max_rids =  max([int(x.split('rId')[-1]) for x in [x['r:id'] for x in out_workbook_sheets]]) if out_workbook_sheets else 0
        b_max_sheets =  len(b_workbook_sheets)
        result = {}
        for element in b_workbook_sheets:
            element['new_sheetId'] = (max_sheets := 1 + max_sheets)
            element['new_rid'] = f"rId{(max_rids := max_rids + 1)}"



        # Actualizo los ids de las hojas de B en xl/workbook.xml
        for element in out_workbook_sheets:
            for sheet in b_workbook_sheets:
                if element['name'] == sheet['name']:
                    new_name = f"{sheet['name']}(1)"
                else:
                    new_name = sheet['name']
                new_sheet_id = sheet['new_sheetId']
                new_rid = sheet['new_rid']

                self.zip_b.files['xl/workbook.xml'] = self.set_workbook_sheet(
                        self.zip_b.get_file_content('xl/workbook.xml'),
                        sheet_index=int(sheet['sheetId']),
                        new_rid=new_rid,
                        new_name=new_name,
                        new_sheet_id=new_sheet_id
                    )
                result.update({sheet['r:id']: new_rid})
        
        # Acutualizo los ids de las hojas de B en xl/_rels/workbook.xml.rels
        for sheet in out_workbook_sheets:

            workbook_rels_xml = self.zip_b.get_file_content('xl/_rels/workbook.xml.rels')
            tree, root = parse_xml_bytes(workbook_rels_xml)

            for _rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rid = _rel.get('Id')
                if result.get(rid):
                    _rel.set('Id', result[rid])
                    self.zip_b.set_file_content('xl/_rels/workbook.xml.rels', write_xml_to_bytes(tree))



    def get_workbook_sheets(self, workbook_content):
        """
        Extrae los r:id, name y sheetId de las etiquetas <sheet> en workbook.xml.

        :param workbook_content: Contenido del archivo workbook.xml en bytes.
        :return: Lista de diccionarios con r:id, name y sheetId.
        """
        # Parsear el contenido de workbook.xml
        root = ET.fromstring(workbook_content)

        # Espacio de nombres utilizado en workbook.xml
        ns = {"": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        # Buscar todas las etiquetas <sheet>
        sheets = root.findall(".//sheet", namespaces=ns)

        # Extraer r:id, name y sheetId
        result = []
        for sheet in sheets:
            result.append({
                "r:id": sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"),
                "name": sheet.attrib.get("name"),
                "sheetId": sheet.attrib.get("sheetId")
            })

        return result

    def set_workbook_sheet(self, workbook_content, sheet_index, new_rid=None, new_name=None, new_sheet_id=None):
        """
        Modifica el r:id, name o sheetId de una hoja específica en workbook.xml.

        :param workbook_content: Contenido del archivo workbook.xml en bytes.
        :param sheet_index: Índice de la hoja a modificar (0 para la primera hoja).
        :param new_rid: Nuevo r:id (opcional).
        :param new_name: Nuevo name (opcional).
        :param new_sheet_id: Nuevo sheetId (opcional).
        :return: Contenido modificado de workbook.xml en bytes.
        """
        # Parsear el contenido de workbook.xml
        root = ET.fromstring(workbook_content)

        # Espacio de nombres utilizado en workbook.xml
        # ns = {"": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        # Buscar todas las etiquetas <sheet>
        sheets = root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet")

        # # Verificar que el índice sea válido
        # if sheet_index < 0 or sheet_index >= len(sheets):
        #     raise IndexError(f"El índice {sheet_index} está fuera de rango. Hay {len(sheets)} hojas.")

        # Obtener la hoja a modificar
        sheet = sheets[sheet_index -1]

        # Modificar los atributos si se proporcionan
        if new_rid:
            sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"] = new_rid
        if new_name:
            sheet.attrib["name"] = new_name
        if new_sheet_id:
            sheet.attrib["sheetId"] = str(new_sheet_id)

        # Convertir el árbol XML de nuevo a bytes
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    
    def merge_elements_deep(self, target, source):
        for src_child in source:
            match = None
            for tgt_child in target:
                if tgt_child.tag == src_child.tag and tgt_child.attrib == src_child.attrib:
                    match = tgt_child
                    break

            if match is not None:
                self.merge_elements_deep(match, src_child)  # fusiona recursivamente
            else:
                target.append(src_child)

    def merge_all_nested_deep(self, xml1, xml2):
        root1 = ET.fromstring(xml1)
        root2 = ET.fromstring(xml2)
        self.merge_elements_deep(root1, root2)
        xml_bytes = ET.tostring(root1, encoding='utf-8', method='xml')
        return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_bytes




    def merge_workbook_strings(self, xml1_str, xml2_str):

        namespaces = {
            'default': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }

        root1 = ET.fromstring(xml1_str)
        root2 = ET.fromstring(xml2_str)

        sheets1 = root1.find('default:sheets', namespaces)
        sheets2 = root2.find('default:sheets', namespaces)

        # Buscar el sheetId más alto que ya tiene workbook1
        existing_ids = [int(sheet.attrib['sheetId']) for sheet in sheets1.findall('default:sheet', namespaces)]
        next_sheet_id = max(existing_ids) + 1 if existing_ids else 1

        # Buscar el rId más alto que ya tiene workbook1
        existing_rids = []
        for sheet in sheets1.findall('default:sheet', namespaces):
            rid = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            if rid.startswith('rId'):
                try:
                    existing_rids.append(int(rid[3:]))
                except ValueError:
                    pass
        next_rid = max(existing_rids) + 1 if existing_rids else 1

        # Agregar cada sheet de workbook2 a workbook1
        for sheet in sheets2.findall('default:sheet', namespaces):
            new_sheet = ET.Element('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet', {
                'name': sheet.attrib['name'],
                'sheetId': str(next_sheet_id),
                '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': f'rId{next_rid}',
            })
            sheets1.append(new_sheet)
            next_sheet_id += 1
            next_rid += 1

        xml_bytes = ET.tostring(root1, encoding='utf-8', method='xml')
        return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_bytes
    

    def merge_relationships_strings(self, xml1_str, xml2_str):

        rel_ns = {
        'default': 'http://schemas.openxmlformats.org/package/2006/relationships'
        }

        root1 = ET.fromstring(xml1_str)
        root2 = ET.fromstring(xml2_str)

        # Buscar todos los rId existentes
        existing_rids = []
        for rel in root1.findall('default:Relationship', rel_ns):
            rid = rel.attrib.get('Id', '')
            if rid.startswith('rId'):
                try:
                    existing_rids.append(int(rid[3:]))
                except ValueError:
                    pass
        next_rid = max(existing_rids) + 1 if existing_rids else 1

        # Agregar cada Relationship del segundo al primero
        for rel in root2.findall('default:Relationship', rel_ns):
            new_rel = ET.Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship', {
                'Id': f'rId{next_rid}',
                'Type': rel.attrib['Type'],
                'Target': rel.attrib['Target'],
            })
            root1.append(new_rel)
            next_rid += 1
        xml_bytes = ET.tostring(root1, encoding='utf-8', method='xml')
        return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_bytes


    def listar_sheets(self, xml_str):
        namespaces = {
                'default': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            }
        root = ET.fromstring(xml_str)
        sheets = root.find('default:sheets', namespaces)
        for sheet in sheets.findall('default:sheet', namespaces):
            name = sheet.attrib['name']
            sheetId = sheet.attrib['sheetId']
            rid = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            print(f"Sheet: name='{name}', sheetId={sheetId}, r:id={rid}")

    def listar_relationships(self, xml_str):
        rel_ns = {
                'default': 'http://schemas.openxmlformats.org/package/2006/relationships'
            }

        root = ET.fromstring(xml_str)
        for rel in root.findall('default:Relationship', rel_ns):
            rid = rel.attrib['Id']
            type_ = rel.attrib['Type']
            target = rel.attrib['Target']
            print(f"Relationship: Id='{rid}', Type='{type_}', Target='{target}'")