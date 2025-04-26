# core/macro_manager.py
import os
import shutil

class MacroManager:
    def __init__(self, output_dir, source_dir):
        self.output_vba_path = os.path.join(output_dir, 'xl', 'vbaProject.bin')
        self.source_vba_path = os.path.join(source_dir, 'xl', 'vbaProject.bin')

    def merge_macros(self):
        if os.path.exists(self.source_vba_path):
            shutil.copyfile(self.source_vba_path, self.output_vba_path)
