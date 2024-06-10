import glob
import os
import sys
from re import A

import openpyxl


class SearchResult:
    def __init__(self, file, sheet, row, col, val, complete_val=None) -> None:
        self.__file = file
        self.__sheet = sheet
        self.__row = row
        self.__col = col
        self.__val = val
        self.__complete_val = complete_val

    @property
    def file(self):
        return self.__file

    @property
    def sheet(self):
        return self.__sheet

    @property
    def cell(self):
        return self.__col + self.__row

    @property
    def value(self):
        return self.__val

    @property
    def is_partial_match(self):
        return self.__complete_val != None


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

    def search_keyword(self, keyword: str):
        exact_match = []
        partial_match = []
        for excel_file in self.excel_files:
            try:
                wb = openpyxl.load_workbook(excel_file, data_only=True)
            except Exception as e:
                print(e)


if __name__ == "__main__":
    recursive = False
    if len(sys.argv) > 1:
        recursive = sys.argv[1].lower() == "true"
    dir_excel = DirectoryExcel(".", recursive)
    print(dir_excel.root_dir)
    for excel_file in dir_excel.excel_files:
        print(excel_file)
