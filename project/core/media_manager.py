# core/media_manager.py
import os
import shutil
from core.rels_manager import RelsManager
from core.xml_utils import generate_unique_id

class MediaManager:
    def __init__(self, output_dir, source_dir):
        self.output_media_dir = os.path.join(output_dir, 'xl', 'media')
        self.source_media_dir = os.path.join(source_dir, 'xl', 'media')
        os.makedirs(self.output_media_dir, exist_ok=True)

    def merge_media(self):
        if not os.path.exists(self.source_media_dir):
            return  # No media to merge

        for filename in os.listdir(self.source_media_dir):
            src_path = os.path.join(self.source_media_dir, filename)
            dst_filename = self._generate_unique_filename(filename)
            dst_path = os.path.join(self.output_media_dir, dst_filename)
            shutil.copyfile(src_path, dst_path)

    def _generate_unique_filename(self, filename):
        name, ext = os.path.splitext(filename)
        existing_files = set(os.listdir(self.output_media_dir))
        i = 1
        new_filename = filename
        while new_filename in existing_files:
            new_filename = f"{name}_{i}{ext}"
            i += 1
        return new_filename
