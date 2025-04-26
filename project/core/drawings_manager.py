# core/drawings_manager.py
import os
import shutil
from core.xml_utils import generate_unique_id
from core.rels_manager import RelsManager

class DrawingsManager:
    def __init__(self, output_dir, source_dir):
        self.output_drawings_dir = os.path.join(output_dir, 'xl', 'drawings')
        self.source_drawings_dir = os.path.join(source_dir, 'xl', 'drawings')
        os.makedirs(self.output_drawings_dir, exist_ok=True)

    def merge_drawings(self):
        if not os.path.exists(self.source_drawings_dir):
            return

        for filename in os.listdir(self.source_drawings_dir):
            src_path = os.path.join(self.source_drawings_dir, filename)
            dst_filename = self._generate_unique_filename(filename)
            dst_path = os.path.join(self.output_drawings_dir, dst_filename)
            shutil.copyfile(src_path, dst_path)

    def _generate_unique_filename(self, filename):
        name, ext = os.path.splitext(filename)
        existing_files = set(os.listdir(self.output_drawings_dir))
        i = 1
        new_filename = filename
        while new_filename in existing_files:
            new_filename = f"{name}_{i}{ext}"
            i += 1
        return new_filename
