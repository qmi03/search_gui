import glob
import os
import sys

import openpyxl


class DirectoryExcel:
    def __init__(self, root_directory: str, recursive: bool = False) -> None:
        self.root_dir = os.path.abspath(root_directory)
        self.recursive = recursive
        self.excel_files = []
        self.find_excel_files()

    def find_excel_files(self):
        extensions = [
            "*.xlsx",
            "*.xls",
            "*.xlsm",
            "*.xlsb",
            "*.xltx",
            "*.xltm",
            "*.xlt",
            "*.xml",
            "*.ods",
        ]
        for extension in extensions:
            if self.recursive:
                search_pattern = os.path.join(self.root_dir, "**", extension)
                self.excel_files.extend(glob.glob(search_pattern, recursive=True))
            else:
                search_pattern = os.path.join(self.root_dir, extension)
                self.excel_files.extend(glob.glob(search_pattern, recursive=False))


dir_excel = DirectoryExcel(".")
print(dir_excel.root_dir)
for excel_file in dir_excel.excel_files:
    print(excel_file)
