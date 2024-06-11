def get_col_letter(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        name = chr(r + ord("A")) + name
    return name
