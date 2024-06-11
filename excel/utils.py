import os


def is_openpyxl_compatible(filename):
    _, extension = os.path.splitext(filename)
    return extension in [".xlsx", ".xlsm", ".xltx", ".xltm"]
