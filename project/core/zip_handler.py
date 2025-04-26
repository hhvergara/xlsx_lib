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
        output = io.BytesIO()
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for path, data in self.files.items():
                zipf.writestr(path, data)
        output.seek(0)
        return output.read()

    def get_file_content(self, path):
        return self.files.get(path)

    def set_file_content(self, path, data):
        self.files[path] = data

    def list_files(self):
        return list(self.files.keys())

    def remove_file(self, path):
        if path in self.files:
            del self.files[path]