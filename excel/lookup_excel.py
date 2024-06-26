import glob
import sys
from pathlib import Path

import pandas as pd

from excel.utils import get_col_letter


class SearchResult:
    def __init__(self, file, sheet, row, col, val, complete_val=None) -> None:
        self.__file = file
        self.__sheet = sheet
        self.__row = row
        self.__col = get_col_letter(col)
        self.__val = val
        self.__complete_val = complete_val if complete_val is not None else val

    def __str__(self):
        return f'Found "{self.found_value}" at cell {self.cell}, sheet {self.sheet}, {Path(self.file).name}'

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
        return self.__complete_val != self.__val

    @property
    def found_value(self):
        return self.__complete_val


class DirectoryExcel:

    def __init__(
        self, root_directory: str = str(Path.home()), recursive: bool = False
    ) -> None:
        self.__root_dir = Path(root_directory).resolve()
        self.__recursive = recursive
        self.excel_files = []
        self.find_excel_files()

    @property
    def root_dir(self):
        return self.__root_dir

    @root_dir.setter
    def root_dir(self, new_root_dir):
        self.excel_files = []
        self.__root_dir = Path(new_root_dir).resolve()
        self.find_excel_files()

    @property
    def recursive(self):
        return self.__recursive

    @recursive.setter
    def recursive(self, new_recursive_policy):
        self.excel_files = []
        self.__recursive = new_recursive_policy
        self.find_excel_files()

    def find_excel_files(self):
        extensions = [
            "*.xlsx",
            "*.xlsm",
            "*.xltx",
            "*.xltm",
            "*.xls",
            "*.xlsb",
            "*.xlt",
            "*.xml",
            "*.ods",
            "*.odt",
        ]
        for extension in extensions:
            files_found = []
            if self.__recursive:
                search_pattern = str(self.__root_dir / "**" / extension)
                files_found = glob.glob(search_pattern, recursive=True)
            else:
                search_pattern = str(self.__root_dir / extension)
                files_found = glob.glob(search_pattern, recursive=False)
            self.excel_files.extend(
                [file for file in files_found if not Path(file).name.startswith("~$")]
            )

    def search_keyword(
        self, keyword, exact_match=False, is_case_sensitive: bool = False
    ):
        if keyword == "":
            return
        for excel_file in self.excel_files:
            try:
                wb = {}
                if Path(excel_file).suffix in [
                    ".xls",
                    ".xlsx",
                    ".xlsm",
                    ".xlsb",
                    ".odf",
                    ".ods",
                    ".odt",
                ]:
                    wb = pd.read_excel(excel_file, sheet_name=None,header=None)
                for sheet_name, ws in wb.items():
                    print(sheet_name)
                    print(ws)
                    for col in ws:
                        print(col)
                        rows = []
                        if (
                            ws[col]
                            .astype(str)
                            .str.contains(keyword, case=is_case_sensitive)
                            .any()
                        ):
                            rows = ws[
                                ws[col]
                                .astype(str)
                                .str.contains(keyword, case=is_case_sensitive)
                            ].index.tolist()
                        for row in rows:
                            yield SearchResult(
                                excel_file,
                                sheet_name,
                                row + 1,
                                col + 1,
                                keyword,
                                ws.loc[row, col],
                            )
            except Exception as e:
                print(f"Error with {excel_file}: {e}")


if __name__ == "__main__":
    recursive = False
    if len(sys.argv) > 1:
        recursive = sys.argv[1].lower() == "true"
    dir_excel = DirectoryExcel(".", recursive)
    print(dir_excel.root_dir)
    for excel_file in dir_excel.excel_files:
        print(excel_file)
