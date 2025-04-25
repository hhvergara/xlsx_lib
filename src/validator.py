import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import tempfile
import os

def validar_relaciones_excel(ruta_xlsx):
    errores = []
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        with zipfile.ZipFile(ruta_xlsx, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        # Buscar todos los r:id en sheetX.xml
        sheet_ids = {}
        sheets_dir = tmpdir / "xl" / "worksheets"
        for sheet_file in sheets_dir.glob("sheet*.xml"):
            try:
                tree = ET.parse(sheet_file)
                for drawing in tree.getroot().findall(".//{*}drawing"):
                    r_id = drawing.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    if r_id:
                        sheet_ids[sheet_file.name] = r_id
            except ET.ParseError:
                errores.append(f"❌ Error parseando {sheet_file.name}")

        # Revisar que esos r:id existan en sheetX.xml.rels
        for sheet_file, r_id in sheet_ids.items():
            rels_file = sheets_dir / "_rels" / f"{sheet_file}.rels"
            if not rels_file.exists():
                errores.append(f"❌ Falta archivo de relaciones: {sheet_file}.rels")
                continue

            try:
                tree = ET.parse(rels_file)
                ids_encontrados = [rel.attrib['Id'] for rel in tree.findall(".//{*}Relationship")]
                if r_id not in ids_encontrados:
                    errores.append(f"❌ {sheet_file} referencia r:id={r_id}, pero no está en su .rels")
            except ET.ParseError:
                errores.append(f"❌ Error parseando relaciones de {sheet_file}")

        # Validar workbook.xml y workbook.xml.rels
        workbook_xml = tmpdir / "xl" / "workbook.xml"
        workbook_rels = tmpdir / "xl" / "_rels" / "workbook.xml.rels"
        if workbook_xml.exists() and workbook_rels.exists():
            try:
                wb_tree = ET.parse(workbook_xml)
                wb_root = wb_tree.getroot()
                wb_rids = [sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                           for sheet in wb_root.findall(".//{*}sheet") if sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")]

                rel_tree = ET.parse(workbook_rels)
                rel_root = rel_tree.getroot()
                rel_ids = [rel.attrib["Id"] for rel in rel_root.findall(".//{*}Relationship")]

                for rid in wb_rids:
                    if rid not in rel_ids:
                        errores.append(f"❌ workbook.xml referencia r:id={rid}, pero no está en workbook.xml.rels")
            except ET.ParseError:
                errores.append("❌ Error parseando workbook.xml o su .rels")

    if errores:
        print("\n🔍 Validación terminada con errores:")
        for err in errores:
            print(err)
    else:
        print("✅ Todas las relaciones r:id están bien referenciadas.")

# Ejemplo de uso
if __name__ == '__main__':
    validar_relaciones_excel("Output.xlsx")
