# main.py
from pathlib import Path
import zipfile
# from core.xlsx_merger import XLSXMerger
from core.merger import XLSXMerger

file_A= 'project/examples/Example A.xlsx'
file_B= 'project/examples/Example B.xlsx'
file_C= 'project/examples/Example C.xlsx' # With images

def descomprimir_xlsx(ruta_xlsx, carpeta_destino):
    Path(carpeta_destino).mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(ruta_xlsx, 'r') as zipf:
        zipf.extractall(carpeta_destino)
    print(f"✅ Extraído en: {carpeta_destino}")


def main():
    # Cargar los archivos de entrada
    with open(file_C, "rb") as f:
        file_a_bytes = f.read()

    with open(file_C, "rb") as f:
        file_b_bytes = f.read()

    # Crear el merger
    merger = XLSXMerger(file_a_bytes, file_b_bytes)

    # Ejecutar el merge
    merged_file_bytes = merger.merge()

    # Guardar el archivo resultante
    with open("merged_output.xlsx", "wb") as f:
        f.write(merged_file_bytes)

    print("Archivo merged_output.xlsx generado exitosamente!")

    descomprimir_xlsx(
            ruta_xlsx="merged_output.xlsx",
            carpeta_destino="merged_output.xlsx".replace('.xlsx', '')
        )
    import shutil
    shutil.copy2("merged_output.xlsx", "merged_output.xlsx".replace('.xlsx', '.zip'))


if __name__ == "__main__":
    main()

