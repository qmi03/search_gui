import glob
import os
import sys
from typing import Generator

import openpyxl
import openpyxl.utils


class SearchResult:
    def __init__(self, file, sheet, row, col, val, complete_val=None) -> None:
        self.__file = file
        self.__sheet = sheet
        self.__row = row
        self.__col = openpyxl.utils.get_column_letter(col)
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
        return self.__col + str(self.__row)

    @property
    def value(self):
        return self.__val

    @property
    def is_partial_match(self):
        return self.__complete_val != None

    @property
    def found_value(self):
        return self.__complete_val


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
            files_found = []
            if self.recursive:
                search_pattern = os.path.join(self.root_dir, "**", extension)
                files_found = glob.glob(search_pattern, recursive=True)
            else:
                search_pattern = os.path.join(self.root_dir, extension)
                files_found = glob.glob(search_pattern, recursive=False)
            self.excel_files.extend([file for file in files_found if not os.path.basename(file).startswith('~$')])

    def search_keyword(
        self, keyword: str, exact_match=False
    ) -> Generator[SearchResult, None, None]:
        for excel_file in self.excel_files:
            try:
                wb = openpyxl.load_workbook(excel_file, data_only=True)
                for ws in wb.worksheets:
                    for col in ws.iter_cols():
                        for cell in col:
                            cell_value = str(cell.value)
                            if keyword == cell_value:
                                yield SearchResult(
                                    excel_file, ws.title, cell.row, cell.column, keyword
                                )
                            elif not exact_match and keyword in cell_value:
                                yield SearchResult(
                                    excel_file,
                                    ws.title,
                                    cell.row,
                                    cell.column,
                                    keyword,
                                    cell_value,
                                )
            except Exception as e:
                print(f"Error opening {excel_file}: {e}")


if __name__ == "__main__":
    recursive = False
    if len(sys.argv) > 1:
        recursive = sys.argv[1].lower() == "true"
    dir_excel = DirectoryExcel(".", recursive)
    print(dir_excel.root_dir)
    for excel_file in dir_excel.excel_files:
        print(excel_file)
