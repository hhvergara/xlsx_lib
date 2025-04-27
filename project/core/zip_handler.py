import zipfile
import io

class ZipHandler:
    def __init__(self, file_bytes=None):
        self.file_bytes = file_bytes
        self.files = {}  # Dict: ruta dentro del zip -> contenido en bytes

    def extract(self):
        with zipfile.ZipFile(io.BytesIO(self.file_bytes), 'r') as zip_ref:
            for zip_info in zip_ref.infolist():
                with zip_ref.open(zip_info.filename) as file:
                    self.files[zip_info.filename] = file.read()

    def create_zip_bytes(self):
        """
        Crea un archivo ZIP en memoria a partir de los archivos almacenados.
        """
        from io import BytesIO
        import zipfile

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_name, file_content in self.files.items():
                zipf.writestr(file_name, file_content)
        zip_buffer.seek(0)
        return zip_buffer.getvalue()

    def get_file_content(self, path):
        return self.files.get(path)

    def set_file_content(self, path, data):
        self.files[path] = data

    def list_files(self):
        return list(self.files.keys())

    def remove_file(self, path):
        if path in self.files:
            del self.files[path]