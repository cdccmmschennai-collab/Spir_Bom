"""
Consolidates multiple already-generated Extraction/OUTPUT files into ONE
shared table per kind (used by the History page's "Combine" action) --
"combine" means merge the actual data rows together, not add a separate
sheet per source file. Selecting N jobs produces at most three sheets
total: "Extraction" (every job's Extraction rows stacked under one
header), "BOM_WORKING" and "SPIR VS TAG" (same idea, from each job's SAP
OUTPUT file) -- never N-per-job tabs. Never re-parses the original SPIR
-- just reads each already-built file's own rows, values and formatting
as-is.

Every file of one kind was produced by the exact same code
(engine.extraction.build_extraction / engine.sap_output.build_sap_output),
so their header row counts are fixed and known here:
  - Extraction:   3 header rows, data starts at row 4  (see build_extraction)
  - BOM_WORKING:  5 header rows, data starts at row 6  (see build_sap_output)
  - SPIR VS TAG:  2 header rows, data starts at row 3  (see build_sap_output)
"""
import os
import tempfile
from copy import copy

import openpyxl

_KINDS = {
    'Extraction': {'header_rows': 3, 'freeze': 'A4'},
    'BOM_WORKING': {'header_rows': 5, 'freeze': 'A6'},
    'SPIR VS TAG': {'header_rows': 2, 'freeze': 'A3'},
}


def _copy_cell(src_cell, dst_ws, row, col):
    dst_cell = dst_ws.cell(row=row, column=col, value=src_cell.value)
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)
    return dst_cell


def _copy_column_widths(src_ws, dst_ws):
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[col_letter].width = dim.width


class _ConsolidatedSheet:
    """One kind's running consolidated sheet -- created lazily on the
    first job that has this kind of data, header copied once from that
    first source, every job's data rows appended underneath in order."""

    def __init__(self, dst_wb, kind: str):
        self.kind = kind
        self.header_rows = _KINDS[kind]['header_rows']
        self.ws = dst_wb.create_sheet(kind)
        self.ws.freeze_panes = _KINDS[kind]['freeze']
        self.next_row = self.header_rows + 1
        self._header_copied = False

    def add(self, src_ws):
        if not self._header_copied:
            for r in range(1, self.header_rows + 1):
                for cell in src_ws[r]:
                    _copy_cell(cell, self.ws, r, cell.column)
            _copy_column_widths(src_ws, self.ws)
            self._header_copied = True

        for r in range(self.header_rows + 1, src_ws.max_row + 1):
            row_cells = list(src_ws[r])
            if all(c.value in (None, '') for c in row_cells):
                continue   # some sheets pad extra blank rows below the real data
            for cell in row_cells:
                _copy_cell(cell, self.ws, self.next_row, cell.column)
            self.next_row += 1


def consolidate_jobs(job_sources) -> str:
    """job_sources: [{'extraction_path': str|None, 'output_path': str|None}, ...],
    one entry per selected job (order preserved). Returns the path to a
    new temp .xlsx containing up to three consolidated sheets -- only the
    kinds that at least one job actually has data for are created at all."""
    dst_wb = openpyxl.Workbook()
    dst_wb.remove(dst_wb.active)
    sheets = {}   # kind -> _ConsolidatedSheet, created on first use

    def _get_sheet(kind):
        if kind not in sheets:
            sheets[kind] = _ConsolidatedSheet(dst_wb, kind)
        return sheets[kind]

    for src in job_sources:
        extraction_path = src.get('extraction_path')
        if extraction_path:
            swb = openpyxl.load_workbook(extraction_path, data_only=True)
            _get_sheet('Extraction').add(swb.active)

        output_path = src.get('output_path')
        if output_path:
            swb = openpyxl.load_workbook(output_path, data_only=True)
            if 'BOM_WORKING' in swb.sheetnames:
                _get_sheet('BOM_WORKING').add(swb['BOM_WORKING'])
            if 'SPIR VS TAG' in swb.sheetnames:
                _get_sheet('SPIR VS TAG').add(swb['SPIR VS TAG'])

    fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    dst_wb.save(tmp_path)
    return tmp_path
