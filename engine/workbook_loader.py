"""
Opens an uploaded SPIR workbook in any supported Excel format as an
openpyxl Workbook, so the parser works the same whatever the format.

.xlsx / .xlsm / .xltx are read by openpyxl directly. .xls (Excel 97-2003,
via xlrd) and .xlsb (binary workbook, via pyxlsb) can't be, so their cell
values are copied -- sheet by sheet, in order, under the same names -- into
a new in-memory openpyxl Workbook. The parser only ever reads cell values,
sheet titles and sheet dimensions (never formatting or merged ranges), so a
values-only copy is all it needs. Converted values are normalized to what
openpyxl itself returns for the same cell in an .xlsx (see _normalize).
"""
import os

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

SUPPORTED_EXTENSIONS = ('.xlsx', '.xlsm', '.xltx', '.xlsb', '.xls')


def load_workbook(path: str):
    ext = os.path.splitext(path)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f'Unsupported file type {ext!r}; expected one of {", ".join(SUPPORTED_EXTENSIONS)}.')
    try:
        if ext == '.xls':
            return _from_xls(path)
        if ext == '.xlsb':
            return _from_xlsb(path)
        return openpyxl.load_workbook(path, data_only=True, keep_vba=ext == '.xlsm')
    except Exception as e:
        raise ValueError(f'Could not read this Excel file ({ext}): it may be corrupt, '
                         f'password-protected, or not really a {ext} file. ({e})') from e


def _normalize(value):
    """Matches openpyxl's .xlsx cell values: blank -> None, and a whole
    number -> int (xlrd/pyxlsb return every number as a float, which would
    otherwise turn e.g. SAP number 10605835 into '10605835.0' downstream).
    Booleans pass through as bool -- the SPIR-type checkbox cells are
    checked with `is True`. Control characters openpyxl can't store in a
    cell (e.g. a stray \\x02 inside a part number) are written as `_x0002_`
    -- exactly how openpyxl reads the same character from an .xlsx -- so a
    SPIR gives the same values, and the same Part Number registry keys,
    whichever format it was uploaded in."""
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub(lambda m: f'_x{ord(m.group()):04X}_', value)
    if value is None or value == '':
        return None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _new_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # drop openpyxl's default empty 'Sheet'
    return wb


def _from_xls(path):
    import xlrd
    book = xlrd.open_workbook(path)
    wb = _new_workbook()
    for sh in book.sheets():
        ws = wb.create_sheet(title=sh.name)
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                if cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    value = bool(cell.value)
                elif cell.ctype == xlrd.XL_CELL_ERROR:   # openpyxl shows these as text too
                    value = xlrd.error_text_from_code.get(cell.value, '#ERROR!')
                elif cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate_as_datetime(cell.value, book.datemode)
                else:
                    value = _normalize(cell.value)
                if value is not None:
                    ws.cell(row=r + 1, column=c + 1, value=value)
    return wb


def _from_xlsb(path):
    import pyxlsb
    wb = _new_workbook()
    with pyxlsb.open_workbook(path) as book:
        for name in book.sheets:
            ws = wb.create_sheet(title=name)
            with book.get_sheet(name) as sh:
                for row in sh.rows(sparse=True):
                    for cell in row:
                        value = _normalize(cell.v)
                        if value is not None:
                            ws.cell(row=cell.r + 1, column=cell.c + 1, value=value)
    return wb
