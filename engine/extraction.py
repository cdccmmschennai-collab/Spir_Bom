"""
Builds the "Extraction" format output: a flat, SAP-field-coded row-per-part
export, grouped tag-major (one tag fully completed -- across every sheet it
appears on -- before the next tag starts).
"""
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.comments import Comment

CENTER = Alignment(horizontal='center', vertical='center')
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                      top=Side(style='thin'), bottom=Side(style='thin'))
HEADER_ROW_FILLS = {
    1: PatternFill(start_color='00B050', end_color='00B050', fill_type='solid'),   # green
    2: PatternFill(start_color='0070C0', end_color='0070C0', fill_type='solid'),   # blue
    3: PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid'),   # red
}

from .rules import (spf_number, spf_prefix, spare_new_description, equipment_new_description,
                    MaterialNumberSeries, equipment_self_item)
from .fx import get_rate_to_qar
from .reference_data import category_for_sap_no

HEADERS = ['S.NO', 'SPIR NO', 'TAG NO', 'EQPT MAKE', 'EQPT MODEL', 'EQPT SR NO', 'EQPT QTY',
           'QUANTITY IDENTICAL PARTS FITTED', 'ITEM NUMBER', 'POSITION NUMBER',
           'OLD MATERIAL NUMBER/SPF NUMBER', 'DESCRIPTION OF PARTS', 'NEW DESCRIPTION OF PARTS',
           'DWG NO INCL POSN NO', 'MANUFACTURER PART NUMBER', 'MATERIAL SPECIFICATION',
           'SUPPLIER/ OCM NAME', 'CURRENCY', 'UNIT PRICE', 'UNIT PRICE (QAR)',
           'DELIVERY TIME IN WEEKS', 'MIN MAX STOCK LVLS QTY', 'UNIT OF MEASURE', 'SAP NUMBER',
           'CLASSIFICATION OF PARTS', 'ERROR', 'SHEET', 'SPIR TYPE', 'VENDOR NAME',
           'VENDOR EMAIL1', 'VENDOR EMAIL2', 'VENDOR CONTACT NO', 'VENDOR COUNTRY']
SAP_CODES = ['NA', 'CDC_SPIRNUMBER', 'EQFNR', 'MFRNR/HERST', 'TYPBZ', 'SERGE', 'MENGE', 'MENGE', 'NA',
             'POSNR', 'BISMT', 'MAKTX', 'MAKTX', 'NA', 'MFPRN', 'NA', 'MFRNR', 'WAERS', 'VERPR', 'VERPR',
             'LEADTIME', 'EISBE', 'MEINS(T006)', 'SAP_MATNR', 'NA', 'ERROR', 'NA', 'NA', 'VENDOR NAME',
             'VENDOR EMAIL1', 'VENDOR EMAIL2', 'VENDOR CONTACT NO', 'VENDOR COUNTRY']
FIELD_LENS = [4, 25, 30, 30, 35, 30, 50, 50, None, 4, 18, 255, 40, 255, 35, 255, 30, 3, 11, 11,
              40, None, 3, 8, 255, 255, 255, 255, 255, 255, 255, 255, 255]
COL_WIDTHS = [6, 26, 15, 12, 12, 18, 10, 14, 10, 10, 22, 40, 45, 22, 25, 18, 16, 10, 10, 12,
              10, 14, 12, 12, 16, 6, 16, 18, 16, 26, 26, 18, 14]


def build_extraction(parsed: dict, out_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = parsed['spir_no'] or 'SPIR'

    bold = Font(bold=True)
    for c, h in enumerate(HEADERS, 1):
        ws.cell(row=1, column=c, value=h).font = bold
    for c, code in enumerate(SAP_CODES, 1):
        ws.cell(row=2, column=c, value=code)
    for c, ln in enumerate(FIELD_LENS, 1):
        if ln is not None:
            ws.cell(row=3, column=c, value=ln)

    sheet_index = {sn: i + 1 for i, sn in enumerate(parsed['sheet_names'])}
    series = MaterialNumberSeries()   # not used for numbering here, extraction keeps SPF-based F col

    row_out = 4
    sno = 0

    def write_row(row: dict):
        nonlocal row_out
        for c, h in enumerate(HEADERS, 1):
            ws.cell(row=row_out, column=c, value=row.get(h, ''))
        row_out += 1

    for tag in parsed['tag_order']:
        info = parsed['tag_info'][tag]
        model, sern = info['model'], info['sern']
        # Vendor details belong to the sheet the record actually came from --
        # never the first sheet's vendor applied blindly to every record.
        origin_sheet = info['sheet']
        equip_vendor = parsed['sheets'][origin_sheet]['vendor']

        sno += 1
        equip_desc_new = equipment_new_description(parsed['equipment_desc'], model, parsed['manufacturer'])
        equip_spf_code = spf_prefix(parsed['spir_no'])

        # If one of this tag's own spare/item rows has the same Mfr Part
        # Number as the equipment's Model Number, that line IS the equipment
        # itself listed as a spare -- mapped 1:1 from the same rule applied
        # to the Output Excel's equivalent columns (SAP Material number,
        # DELIVERY TIME IN WEEKS, MM REQUIREMENT MOVING AVERAGE PRICE,
        # CURRENCY CODE REF -> SAP NUMBER, DELIVERY TIME IN WEEKS, UNIT
        # PRICE, CURRENCY here). UNIT PRICE stays the self item's own raw,
        # native-currency price -- consistent with what UNIT PRICE already
        # means for every other row in this file (Output's version is
        # QAR-converted only because its own price column always is).
        self_item = equipment_self_item(parsed, tag, model)
        equip_sap_no = equip_delivery = equip_unit_price = equip_currency = None
        if self_item is not None:
            equip_sap_no = self_item['sap_no']
            equip_delivery = self_item['delivery_wks']
            equip_unit_price = self_item['unit_price']
            equip_currency = self_item['currency']

        write_row({
            'S.NO': f'{sno:04d}', 'SPIR NO': parsed['spir_no'], 'TAG NO': tag,
            'EQPT MAKE': parsed['manufacturer'], 'EQPT MODEL': model, 'EQPT SR NO': sern,
            'EQPT QTY': 1, 'POSITION NUMBER': '0010',
            'OLD MATERIAL NUMBER/SPF NUMBER': equip_spf_code,
            'DESCRIPTION OF PARTS': parsed['equipment_desc'], 'NEW DESCRIPTION OF PARTS': equip_desc_new,
            'MANUFACTURER PART NUMBER': model, 'SUPPLIER/ OCM NAME': parsed['supplier'],
            'CURRENCY': equip_currency, 'UNIT PRICE': equip_unit_price,
            'DELIVERY TIME IN WEEKS': equip_delivery, 'UNIT OF MEASURE': 'EA',
            'SAP NUMBER': equip_sap_no, 'CLASSIFICATION OF PARTS': 'B',
            'ERROR': 0, 'SHEET': origin_sheet, 'SPIR TYPE': parsed['spir_type'],
            'VENDOR NAME': equip_vendor['name'],
            'VENDOR EMAIL1': equip_vendor['email1'], 'VENDOR EMAIL2': equip_vendor['email2'],
            'VENDOR CONTACT NO': equip_vendor['contact_no'], 'VENDOR COUNTRY': equip_vendor['country'],
        })

        pos_counter = 0   # restarts at 0010 for EACH equipment/tag
        for sn in parsed['sheet_names']:
            sheet = parsed['sheets'][sn]
            for it in sheet['items']:
                if tag not in it['flags']:
                    continue
                sno += 1
                pos_counter += 1
                pos_no = f'{pos_counter * 10:04d}'
                spf_code = spf_number(parsed['spir_no'], sheet_index[sn], it['item_no'])
                new_desc = spare_new_description(it)
                rate = get_rate_to_qar(it['currency'])
                qar_price = round(it['unit_price'] * rate, 2) if isinstance(it['unit_price'], (int, float)) and rate else ''
                # Same fix as the Output file's Quantity column: the real
                # per-BOM-line quantity is this tag's own value on its flag
                # cell (C:F) for this item row, not 'qty_fitted' (column H,
                # "TOTAL NO. OF IDENTICAL PARTS FITTED" -- an equipment-level
                # count, e.g. 16 parts fitted, not the 1 actually ordered).
                bom_qty = it['tag_qty'].get(tag)
                write_row({
                    'S.NO': f'{sno:04d}', 'SPIR NO': sheet['spir_no'], 'TAG NO': tag,
                    'EQPT MAKE': sheet['manufacturer'], 'EQPT MODEL': model, 'EQPT SR NO': sern,
                    'QUANTITY IDENTICAL PARTS FITTED': bom_qty, 'ITEM NUMBER': it['item_no'],
                    'POSITION NUMBER': pos_no, 'OLD MATERIAL NUMBER/SPF NUMBER': spf_code,
                    'DESCRIPTION OF PARTS': it['desc'], 'NEW DESCRIPTION OF PARTS': new_desc,
                    'DWG NO INCL POSN NO': it['dwg_no'], 'MANUFACTURER PART NUMBER': it['mfr_part_no'],
                    'MATERIAL SPECIFICATION': it['material_spec'], 'SUPPLIER/ OCM NAME': it['supplier_ocm'],
                    'CURRENCY': it['currency'], 'UNIT PRICE': it['unit_price'], 'UNIT PRICE (QAR)': qar_price,
                    'DELIVERY TIME IN WEEKS': it['delivery_wks'], 'UNIT OF MEASURE': it['uom'],
                    'SAP NUMBER': it['sap_no'], 'CLASSIFICATION OF PARTS': category_for_sap_no(it['sap_no']),
                    'ERROR': 0, 'SHEET': sn, 'SPIR TYPE': sheet['spir_type'],
                    'VENDOR NAME': sheet['vendor']['name'],
                    'VENDOR EMAIL1': sheet['vendor']['email1'], 'VENDOR EMAIL2': sheet['vendor']['email2'],
                    'VENDOR CONTACT NO': sheet['vendor']['contact_no'], 'VENDOR COUNTRY': sheet['vendor']['country'],
                })

    for i, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    ws.freeze_panes = 'A4'   # keep rows 1-3 visible while scrolling
    last_row = row_out - 1
    for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=len(HEADERS)):
        for cell in row:
            cell.alignment = CENTER
            cell.border = THIN_BORDER
            if cell.row in HEADER_ROW_FILLS:
                cell.fill = HEADER_ROW_FILLS[cell.row]

    ws['T1'].comment = Comment(
        "UNIT PRICE (QAR) converted using a live FX rate fetched at generation time "
        "(falls back to the official USD/QAR peg if the live lookup fails).", "BOM Tool")

    wb.save(out_path)
    return out_path
