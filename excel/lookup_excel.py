import glob
import os
import sys
from re import A

import openpyxl


class DirectoryExcel:
    class SearchResult:
        def __init__(self,file,sheet,row,col,val) -> None:
            self.file = file
            self.sheet = sheet
            self.row = row
            self.col = col
            self.val = val
        def __str__(self):
            return f"File: {self.file}, Sheet: {self.sheet}, Row: {self.row}, Column: {self.col}, Value: {self.val}"

        def get_file(self):
            return self.file

        def get_sheet(self):
            return self.sheet

        def get_row(self):
            return self.row

        def get_column(self):
            return self.col

        def get_value(self):
            return self.val

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
    def search_keyword (self,keyword:str):




if __name__ == "__main__":
    recursive = False
    if len(sys.argv) > 1:
        recursive = sys.argv[1].lower() == "true"
    dir_excel = DirectoryExcel(".", recursive)
    print(dir_excel.root_dir)
    for excel_file in dir_excel.excel_files:
        print(excel_file)
