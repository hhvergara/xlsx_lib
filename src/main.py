import os
import zipfile
import shutil
import tempfile
from pathlib import Path
import xml.etree.ElementTree as ET


file_1 = "Example.xlsx"
file_2 = "Example B.xlsx"

# Eliminar prefijos innecesarios
ET.register_namespace('', "http://schemas.openxmlformats.org/package/2006/relationships")
ET.register_namespace('r', "http://schemas.openxmlformats.org/officeDocument/2006/relationships")



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




def obtener_proximo_rid(rels_path, r):
    tree = ET.parse(rels_path)
    root = tree.getroot()
    ids = [int(rel.attrib["Id"][3:]) for rel in root.findall(f"rel:Relationship", r)
           if rel.attrib["Id"].startswith("rId") and rel.attrib["Id"][3:].isdigit()]
    return f"rId{max(ids) + 1}" if ids else "rId1"

def copiar_hoja_completa_con_imagenes(origen_xlsx, destino_xlsx, nombre_hoja='Sheet1', output_filename='output.xlsx'):
    r = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
        "ct": "http://schemas.openxmlformats.org/package/2006/content-types"
    }

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        origen_dir = temp_dir / "origen"
        destino_dir = temp_dir / "destino"

        with zipfile.ZipFile(origen_xlsx, "r") as zip_ref:
            zip_ref.extractall(origen_dir)

        with zipfile.ZipFile(destino_xlsx, "r") as zip_ref:
            zip_ref.extractall(destino_dir)

        worksheets_dir = destino_dir / "xl" / "worksheets"
        drawings_dir = destino_dir / "xl" / "drawings"
        media_dir = destino_dir / "xl" / "media"
        rels_dir = destino_dir / "xl" / "_rels"

        existing_sheets = list(worksheets_dir.glob("sheet*.xml"))
        next_sheet_num = max([int(f.stem.replace("sheet", "")) for f in existing_sheets]) + 1
        nueva_hoja_filename = f"sheet{next_sheet_num}.xml"

        existing_drawings = list(drawings_dir.glob("drawing*.xml"))
        next_drawing_num = max([int(f.stem.replace("drawing", "")) for f in existing_drawings], default=0) + 1
        nueva_drawing_filename = f"drawing{next_drawing_num}.xml"

        shutil.copyfile(origen_dir / "xl" / "worksheets" / "sheet1.xml",
                        worksheets_dir / nueva_hoja_filename)

        origen_sheet_rels = origen_dir / "xl" / "worksheets" / "_rels" / "sheet1.xml.rels"
        if origen_sheet_rels.exists():
            dest_sheet_rels_dir = worksheets_dir / "_rels"
            dest_sheet_rels_dir.mkdir(exist_ok=True)
            nueva_sheet_rels_filename = f"{nueva_hoja_filename}.rels"
            shutil.copyfile(origen_sheet_rels, dest_sheet_rels_dir / nueva_sheet_rels_filename)

            tree = ET.parse(dest_sheet_rels_dir / nueva_sheet_rels_filename)
            root = tree.getroot()
            for rel in root.findall("rel:Relationship", r):
                if "drawing" in rel.attrib["Target"]:
                    rel.attrib["Target"] = f"../drawings/{nueva_drawing_filename}"
            tree.write(dest_sheet_rels_dir / nueva_sheet_rels_filename, encoding="utf-8", xml_declaration=True)

        origen_drawing = origen_dir / "xl" / "drawings" / "drawing1.xml"
        if origen_drawing.exists():
            shutil.copyfile(origen_drawing, drawings_dir / nueva_drawing_filename)

        origen_drawing_rels = origen_dir / "xl" / "drawings" / "_rels" / "drawing1.xml.rels"
        if origen_drawing_rels.exists():
            dest_drawing_rels_dir = drawings_dir / "_rels"
            dest_drawing_rels_dir.mkdir(exist_ok=True)
            nueva_drawing_rels_filename = f"{nueva_drawing_filename}.rels"
            shutil.copyfile(origen_drawing_rels, dest_drawing_rels_dir / nueva_drawing_rels_filename)

            tree = ET.parse(dest_drawing_rels_dir / nueva_drawing_rels_filename)
            root = tree.getroot()
            for rel in root.findall("rel:Relationship", r):
                if "media" in rel.attrib["Target"]:
                    img_name = os.path.basename(rel.attrib["Target"])
                    origen_img = origen_dir / "xl" / "media" / img_name
                    if origen_img.exists():
                        shutil.copyfile(origen_img, media_dir / img_name)
            tree.write(dest_drawing_rels_dir / nueva_drawing_rels_filename, encoding="utf-8", xml_declaration=True)

        workbook_path = destino_dir / "xl" / "workbook.xml"
        tree = ET.parse(workbook_path)
        root = tree.getroot()
        sheets = root.find("main:sheets", r)

        rels_path = rels_dir / "workbook.xml.rels"
        nuevo_rid = obtener_proximo_rid(rels_path, r)

        new_sheet = ET.Element("sheet", {
            "name": nombre_hoja,
            "sheetId": str(next_sheet_num),
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": nuevo_rid
        })
        sheets.append(new_sheet)
        tree.write(workbook_path, encoding="UTF-8", xml_declaration=True)

        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        new_rel = ET.Element("Relationship", {
            "Id": nuevo_rid,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "Target": f"worksheets/{nueva_hoja_filename}"
        })
        rels_root.append(new_rel)
        rels_tree.write(rels_path, encoding="UTF-8", xml_declaration=True)

        ct_path = destino_dir / "[Content_Types].xml"
        ct_tree = ET.parse(ct_path)
        ct_root = ct_tree.getroot()

        ET.SubElement(ct_root, "Override", {
            "PartName": f"/xl/worksheets/{nueva_hoja_filename}",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        })

        if origen_drawing.exists():
            ET.SubElement(ct_root, "Override", {
                "PartName": f"/xl/drawings/{nueva_drawing_filename}",
                "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml"
            })

        for img_file in (origen_dir / "xl" / "media").glob("*"):
            img_ext = img_file.suffix.lower().replace(".", "")
            mime_type = {
                "png": "image/png",
                "jpeg": "image/jpeg",
                "jpg": "image/jpeg",
                "gif": "image/gif"
            }.get(img_ext)
            if mime_type:
                override = f"/xl/media/{img_file.name}"
                if not any(o.attrib.get("PartName") == override for o in ct_root.findall("Override")):
                    ET.SubElement(ct_root, "Override", {
                        "PartName": override,
                        "ContentType": mime_type
                    })

        ct_tree.write(ct_path, encoding="UTF-8", xml_declaration=True)

        nuevo_destino = output_filename
        # nuevo_destino = destino_xlsx.replace(".xlsx", " convertido.xlsx")
        with zipfile.ZipFile(nuevo_destino, "w", zipfile.ZIP_DEFLATED) as zipf:
            for foldername, subfolders, filenames in os.walk(destino_dir):
                for filename in filenames:
                    file_path = Path(foldername) / filename
                    arcname = file_path.relative_to(destino_dir)
                    zipf.write(file_path, arcname)

        print(f"✅ Hoja copiada con imágenes exitosamente a: {nuevo_destino}")
    return output_filename


if __name__ == "__main__":
    output = copiar_hoja_completa_con_imagenes(file_1, file_2)
    # Ejemplo de uso
    descomprimir_xlsx(output, output.replace(".xlsx", "C"))
    # Ejemplo de uso
    # recomprimir_a_xlsx('Example 2 tabs', 'Example 2 tabs')