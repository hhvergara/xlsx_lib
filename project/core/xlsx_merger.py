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

        # Inspección: Lista los archivos que se incluirán en el ZIP
        print("Archivos en el ZIP resultante:")
        for file_name in self.output_zip.files.keys():
            print(f" - {file_name}")

        return self.output_zip.create_zip_bytes()

    def merge_sheets(self):
        """
        Fusiona las hojas de cálculo y recursos de los archivos A y B en el archivo Output.
        Asegura que los identificadores y nombres sean únicos y que las relaciones sean válidas.
        """
        # 1. Cargar el contenido de A en RAM y copiarlo al archivo Output
        content_a = self.zip_a.files
        self.output_zip.files = dict(content_a)  # Copiar todo el contenido de A al Output

        # 2. Cargar el contenido de B en RAM
        content_b = self.zip_b.files

        # 3. Procesar las hojas de cálculo de B
        for file_name, file_content in content_b.items():
            if file_name.startswith("xl/worksheets/sheet"):  # Detectar hojas de cálculo
                self._merge_sheet(file_name, file_content)

            elif file_name.startswith("xl/media/"):  # Detectar imágenes
                self._merge_media(file_name, file_content)

            elif file_name.startswith("xl/charts/"):  # Detectar gráficos
                self._merge_chart(file_name, file_content)

            else:
                # Otros recursos (e.g., estilos, sharedStrings)
                self._merge_generic_resource(file_name, file_content)

        # 4. Actualizar los archivos de relaciones y metadatos
        self._update_relationships()
        self._update_docProps()

    def _merge_sheet(self, file_name, file_content):
        """
        Fusiona una hoja de cálculo, asegurando que el nombre y los IDs sean únicos.
        """
        # Generar un nuevo nombre único para la hoja
        existing_sheets = [sheet.attrib['name'] for sheet in self._get_sheets()]
        sheet_name = file_name.split("/")[-1].replace(".xml", "")
        new_sheet_name = sheet_name
        counter = 1
        while new_sheet_name in existing_sheets:
            new_sheet_name = f"{sheet_name}({counter})"
            counter += 1

        # Generar un nuevo ID único
        existing_ids = {int(sheet.attrib['sheetId']) for sheet in self._get_sheets()}
        new_sheet_id = max(existing_ids) + 1

        # Actualizar el contenido de la hoja con el nuevo nombre y ID
        tree, root = parse_xml_bytes(file_content)
        root.attrib['name'] = new_sheet_name
        root.attrib['sheetId'] = str(new_sheet_id)

        # Guardar la hoja en el archivo Output
        new_file_name = f"xl/worksheets/sheet{new_sheet_id}.xml"
        self.output_zip.set_file_content(new_file_name, write_xml_to_bytes(tree))

        # Agregar la hoja al workbook.xml
        self._add_sheet_to_workbook(new_sheet_name, new_sheet_id, new_file_name)

    def _merge_media(self, file_name, file_content):
        """
        Fusiona una imagen, asegurando que el nombre sea único.
        """
        existing_media = [name for name in self.output_zip.files if name.startswith("xl/media/")]
        new_file_name = file_name
        counter = 1
        while new_file_name in existing_media:
            base, ext = file_name.rsplit(".", 1)
            new_file_name = f"{base}({counter}).{ext}"
            counter += 1

        # Guardar la imagen en el archivo Output
        self.output_zip.set_file_content(new_file_name, file_content)

    def _merge_chart(self, file_name, file_content):
        """
        Fusiona un gráfico, asegurando que el nombre sea único.
        """
        existing_charts = [name for name in self.output_zip.files if name.startswith("xl/charts/")]
        new_file_name = file_name
        counter = 1
        while new_file_name in existing_charts:
            base, ext = file_name.rsplit(".", 1)
            new_file_name = f"{base}({counter}).{ext}"
            counter += 1

        # Guardar el gráfico en el archivo Output
        self.output_zip.set_file_content(new_file_name, file_content)

    def _merge_generic_resource(self, file_name, file_content):
        """
        Fusiona recursos genéricos (e.g., estilos, sharedStrings).
        Si ya existe en el archivo Output, no lo sobrescribe.
        """
        if file_name not in self.output_zip.files:
            self.output_zip.set_file_content(file_name, file_content)

    def _update_relationships(self):
        """
        Actualiza los archivos de relaciones (_rels) para reflejar los cambios en los recursos.
        """
        # Actualizar workbook.xml.rels
        workbook_rels = self.output_zip.get_file_content("xl/_rels/workbook.xml.rels")
        tree, root = parse_xml_bytes(workbook_rels)

        # Agregar relaciones para las nuevas hojas, imágenes y gráficos
        for file_name in self.output_zip.files:
            if file_name.startswith("xl/worksheets/"):
                self._add_relationship(root, file_name, "worksheet")
            elif file_name.startswith("xl/media/"):
                self._add_relationship(root, file_name, "image")
            elif file_name.startswith("xl/charts/"):
                self._add_relationship(root, file_name, "chart")

        # Guardar los cambios
        self.output_zip.set_file_content("xl/_rels/workbook.xml.rels", write_xml_to_bytes(tree))

    def _update_docProps(self):
        """
        Actualiza los archivos de metadatos (docProps) para reflejar los cambios.
        """
        # Actualizar app.xml
        self.update_app_xml()

        # Actualizar core.xml
        self.update_core_xml()

    def _add_sheet_to_workbook(self, sheet_name, sheet_id, file_name):
        """
        Agrega una hoja al archivo workbook.xml.
        """
        workbook = self.output_zip.get_file_content("xl/workbook.xml")
        tree, root = parse_xml_bytes(workbook)
        sheets = root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets")

        # Crear un nuevo elemento para la hoja
        new_sheet = ET.Element("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet", {
            "name": sheet_name,
            "sheetId": str(sheet_id),
            "r:id": f"rId{sheet_id}"
        })
        sheets.append(new_sheet)

        # Guardar los cambios
        self.output_zip.set_file_content("xl/workbook.xml", write_xml_to_bytes(tree))

    def _add_relationship(self, root, target, type_):
        """
        Agrega una relación al archivo de relaciones.
        """
        type_map = {
            "worksheet": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "image": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            "chart": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
        }
        new_id = generate_unique_id({rel.attrib["Id"] for rel in root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")})
        ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship", {
            "Id": new_id,
            "Type": type_map[type_],
            "Target": target
        })

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

        # Asegurar que todas las hojas estén registradas
        content_types = self.output_zip.get_file_content("[Content_Types].xml")
        tree, root = parse_xml_bytes(content_types)
        for sheet_id in range(1, len(self.output_zip.files) + 1):
            part_name = f"/xl/worksheets/sheet{sheet_id}.xml"
            if not any(override.attrib['PartName'] == part_name for override in root.findall('{http://schemas.openxmlformats.org/package/2006/content-types}Override')):
                ET.SubElement(root, "Override", {
                    "PartName": part_name,
                    "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                })
        self.output_zip.set_file_content("[Content_Types].xml", write_xml_to_bytes(tree))

    def find_target_in_rels(self, rels_bytes, r_id):
        """
        Encuentra el destino de una relación dado su ID.
        Maneja relaciones internas y externas.
        """
        tree, root = parse_xml_bytes(rels_bytes)
        for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            if rel.attrib['Id'] == r_id:
                target = rel.attrib['Target']
                target_mode = rel.attrib.get('TargetMode', 'Internal')
                return target, target_mode
        return None, None

    def get_existing_rids(self, rels_bytes):
        tree, root = parse_xml_bytes(rels_bytes)
        return {rel.attrib['Id'] for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')}

    def append_to_rels(self, rels_bytes, new_id, type_, target, target_mode=None):
        """
        Agrega una nueva relación al archivo de relaciones.

        :param rels_bytes: Contenido del archivo de relaciones en bytes.
        :param new_id: Identificador único para la relación (e.g., rId1).
        :param type_: Tipo de la relación (e.g., worksheet, hyperlink).
        :param target: URI del destino de la relación.
        :param target_mode: Modo del destino (e.g., 'External' para enlaces externos).
        """
        tree, root = parse_xml_bytes(rels_bytes)
        attributes = {'Id': new_id, 'Type': type_, 'Target': target}
        if target_mode:
            attributes['TargetMode'] = target_mode
        ET.SubElement(root, 'Relationship', attributes)
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

    def _get_sheets(self):
        """
        Obtiene todas las hojas de cálculo del archivo workbook.xml.
        :return: Lista de elementos XML que representan las hojas.
        """
        workbook = self.output_zip.get_file_content("xl/workbook.xml")
        tree, root = parse_xml_bytes(workbook)
        sheets = root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets")
        return sheets.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet")
