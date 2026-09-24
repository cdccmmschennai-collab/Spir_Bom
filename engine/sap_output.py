"""
Builds the "OUTPUT" format: the 41-column SAP material-master upload
structure (Equipment/BOM-header rows + spare/leaf rows), matching the
reference file's header banding and duplicate-highlight conditional
formatting.

Every reference-driven column (country, material group, material type,
material category, plant, description abbreviation) is sourced from
Reference.xlsx via engine.reference_data -- nothing here is hard-coded.
Anything that can't be determined confidently is left blank and appended to
data/output_lookup_review.log for manual follow-up.
"""
import os
import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle

from .rules import (spare_new_description, equipment_new_description, spf_number,
                    equipment_dedup_key, spir_type_short_code, equipment_self_item)
from .part_master import PersistentMaterialNumberSeries
from .fx import get_rate_to_qar
from .manufacturer_country import resolve_country
from . import reference_data
from .db import DATA_DIR
from . import db

REVIEW_LOG_PATH = os.path.join(DATA_DIR, 'output_lookup_review.log')

HEADERS = ['Serial Number', 'Tag Number', ' SAP Material number', 'Material Temp Number', 'Material TypeVCategory',
           'External Number Assigned for Assembly', 'Old Material Number/SPF Number', 'MESC/SPF_Ref',
           'LEN(OLD MAT/SPF NUMBER = 18)', 'NEW DESCRIPTION OF PARTS', 'Material Description', 'LEN(MAT.DESC = 40)',
           'MAINTENANCE PLANNING PLANT', 'Base Unit of Measure(T006) ', 'Manufacturers Part Number',
           'LEN(MFR PART NUMBER = 18)', 'Manufacturer Name', 'LEN(Manufacturer=30)', 'Safety Stock',
           'Criticality (Equipment)', 'Material is an Equipment Indicator ', 'BOM HEADER  Number  ',
           'Quantity', 'Position Number', 'CDC_SPIR NUMBER', 'DELIVERY TIME IN WEEKS',
           'Manufacturer Country CODE', 'MM REQUIREMENT MATERIAL TYPE', 'MM REQUIREMENT MATERIAL GROUP ',
           'MM REQUIREMENT MOVING AVERAGE PRICE', 'PO TEXT', 'CDC_ADD INFORMATION',
           'Manufacturer COUNTRY NAME', 'CURRENCY CODE REF', 'BOM HEADER-REMARKS', 'PRD NAME',
           'PRD DATE \nDD-MM-YYYY', 'SELF QC', 'LEAD QC  NAME', 'LEAD QC DATE \nDD-MM-YYYY',
           'TEAM LEAD QC', 'DATE']
SAP_CODES = ['NA', 'EQFNR', 'SAP_MATNR', 'MATNR', 'POSTP', 'EXTMATNR', 'BISMT', None, 'NA', 'MAKTX', 'MAKTX', 'NA',
             'IWERK', 'MEINS', 'MFPRN', 'NA', 'MFRNR', 'NA', 'EISBE', 'ABC', 'SUBMT', 'BOM_PART_EQUIP',
             'MENGE', 'POSNR', 'CDC_SPIR NUMBER', 'LEAD TIME ', 'HERLD', 'MTART', 'MATKL', 'VERPR',
             'PO TEXT', 'CDC_ADD INFO', 'COUNTRY NAME', 'CURRENCY', 'REMARKS', 'NAME', 'DD-MM-YYYY',
             'SELF QC', 'NAME', 'DD-MM-YYYY', 'TEAM LEAD QC', None]
FIELD_LENS = [4, 30, 8, 8, 1, 4, 18, None, 18, 40, 40, 40, 4, 3, 35, 35, 30, 30, 50, 1, 3, 8, 50, 4, 25, 40,
              3, 4, 4, 11, 255, 255, 255, 3, 255, 255, 10, 255, 255, 10, 255, None]
ROW4 = ['NA', 'NA', 'NA', 'NA', 'T418', 'NA', 'NA', None, 'NA', 'NA', 'NA', 'NA', 'T399I', 'T006 ',
        'REFERENCE-ANEXURE1-MANUFACTURER PART NUMBER', 'NA', 'REFERENCE-ANEXURE1-MANUFACTURER NAME',
        'NA', 'NA', 'NA', 'YES/NO', 'NA', 'NA', 'NA', 'NA', 'NA', 'T005', 'T134', ' T023', 'NA', 'NA',
        'NA', 'NA', 'NA', 'NA', 'REFERENCE_NAME', 'NA', 'COMPLETED /INPROGRESS', 'REFERENCE_NAME',
        'NA', 'COMPLETED /INPROGRESS', None]

# HEADERS unpacked by position into named constants -- used as write_row dict
# keys so a column is always addressed by its guaranteed-correct position in
# HEADERS, never by retyping a header string (several have leading/double/
# trailing spaces that are easy to get subtly wrong).
(H_SNO, H_TAG, H_SAP_MATNR, H_MAT_TEMP, H_CAT, H_EXT_NUM, H_OLD_MATNR, H_MESC, H_LEN_OLD,
 H_NEW_DESC, H_MAT_DESC, H_LEN_DESC, H_PLANT, H_UOM, H_MFR_PARTNO, H_LEN_MFR, H_MFR_NAME,
 H_LEN_MFR_NAME, H_SAFETY, H_CRIT, H_IS_EQUIP, H_BOM_HDR_NUM, H_QTY, H_POS_NUM, H_CDC_SPIR,
 H_DELIVERY, H_COUNTRY_CODE, H_MM_TYPE, H_MM_GROUP, H_MM_PRICE, H_PO_TEXT, H_ADD_INFO,
 H_COUNTRY_NAME, H_CURRENCY, H_REMARKS, H_PRD_NAME, H_PRD_DATE, H_SELF_QC, H_LEAD_QC_NAME,
 H_LEAD_QC_DATE, H_TEAM_LEAD, H_DATE) = HEADERS

ROW_FILLS = {
    1: {'default': 'FFFF00', 1: '00B050'},   # yellow, col A green
    2: {'default': '002060'},                 # navy
    3: {'default': 'C00000'},                 # red
    4: {'default': '7030A0'},                 # purple
    5: {'default': '00B0F0'},                 # light blue
}
DUP_FILL = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
DUP_FONT = Font(color='9C0006')

CENTER_WRAP = Alignment(horizontal='center', vertical='center', wrap_text=True)
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                      top=Side(style='thin'), bottom=Side(style='thin'))
GREEN_FILL = PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')

# 'SPIR VS TAG' sheet: two fixed header rows only (no data populated yet).
# Row 1 is the SAP field-code row, row 2 the human-readable label row --
# same two-row convention as BOM_WORKING, laid out exactly as specified.
SPIR_VS_TAG_ROW1 = ['EQFNR', 'SUBMT', 'SPIR NO', 'REV', 'HERST', 'TYPBZ', 'SERGE',
                    'REMARKS', 'DATE', 'FMTL MATCH', 'SPIR TYPE', 'TAGAS PER SPIR DOC', 'TYPE']
SPIR_VS_TAG_ROW2 = ['TAG NO', 'material Construction type', 'SPIR NO', 'REV',
                    'MANUFACTURER NAME', 'MODEL', 'SEREAL NO', 'REMARKS', 'DATE',
                    'FMTL MATCH', 'SPIR TYPE', '', '']


def _apply_header_styling(ws, ncols=42):
    white_bold = Font(bold=True, color='FFFFFF')
    black_bold = Font(bold=True, color='000000')
    black = Font(color='000000')
    for row, spec in ROW_FILLS.items():
        for c in range(1, ncols + 1):
            color = spec.get(c, spec['default'])
            fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
            cell = ws.cell(row=row, column=c)
            cell.fill = fill
            cell.border = THIN_BORDER
            if row <= 4:
                cell.font = white_bold if row == 2 else black_bold
            else:
                cell.font = black
            cell.alignment = CENTER_WRAP


def _build_spir_vs_tag_sheet(wb, duplicate_rows):
    """Two styled/frozen header rows (unchanged, exact field names), plus one
    data row per duplicate TAG (see build_sap_output's dedup pass)."""
    ws = wb.create_sheet('SPIR VS TAG')
    header_font = Font(bold=True, color='000000')
    for r, row_vals in enumerate((SPIR_VS_TAG_ROW1, SPIR_VS_TAG_ROW2), start=1):
        for c, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r, column=c, value=val or None)
            cell.fill = GREEN_FILL
            cell.font = header_font
            cell.alignment = CENTER_WRAP
            cell.border = THIN_BORDER
    for c in range(1, len(SPIR_VS_TAG_ROW1) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 18
    for r, row_vals in enumerate(duplicate_rows, start=3):
        for c, val in enumerate(row_vals, start=1):
            ws.cell(row=r, column=c, value=val)
    ws.freeze_panes = 'A3'
    return ws


def _first_item_mfr_part_for_tag(parsed, tag):
    """The Manufacturer Part Number of the first spare/item row flagged for
    this tag (sheet order, then row order) -- used as the equipment-level
    'Manufacturer Part Number' half of the duplicate-detection key, since
    the SPIR itself has no such field at the equipment/tag level."""
    for sn in parsed['sheet_names']:
        for it in parsed['sheets'][sn]['items']:
            if tag in it['flags']:
                return it['mfr_part_no']
    return None


def _qar_price(unit_price, currency_raw):
    """(qar_price, currency_code) for a raw unit price + the SPIR's full
    currency dropdown text (e.g. 'USD - United States Dollar' -> 'USD')."""
    currency_code = str(currency_raw or '').split(' ')[0]
    rate = get_rate_to_qar(currency_code) if currency_code else None
    qar_price = round(unit_price * rate, 2) if isinstance(unit_price, (int, float)) and rate else None
    return qar_price, currency_code


def _find_duplicate_tags(parsed):
    """Groups parsed['tag_order'] by (Model Number, first item's Mfr Part
    Number). Returns duplicate_of: {tag -> first_tag} for every TAG that is
    NOT the first one seen for its (model, mfr_part_no) combination. A tag
    with a missing model or mfr_part_no is never treated as a duplicate."""
    group_first_tag = {}
    duplicate_of = {}
    for tag in parsed['tag_order']:
        model = parsed['tag_info'][tag]['model']
        first_mfr = _first_item_mfr_part_for_tag(parsed, tag)
        key = equipment_dedup_key(model, first_mfr)
        if key is None:
            continue
        if key not in group_first_tag:
            group_first_tag[key] = tag
        elif tag not in duplicate_of:
            duplicate_of[tag] = group_first_tag[key]
    return duplicate_of


def _log_review(spir_filename: str, tag: str, manufacturer: str, column: str, reason: str):
    os.makedirs(DATA_DIR, exist_ok=True)
    try:
        with open(REVIEW_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(f'{datetime.datetime.utcnow().isoformat()}\t{spir_filename}\t{tag}\t'
                    f'{manufacturer}\t{column}\t{reason}\n')
    except Exception:
        pass


def _part_output_row(get):
    """The Parts Master-searchable attributes of one BOM_WORKING row (see
    db.part_output_rows), given `get(header) -> cell value`. Shared by
    build_sap_output and the backfill below so both read the same columns."""
    def text(header):
        v = get(header)
        return None if v in (None, '') else str(v).strip()
    is_equip = (text(H_IS_EQUIP) or '').upper()
    return {
        'material_temp_number': get(H_MAT_TEMP),
        'part_number': text(H_MFR_PARTNO),
        'part_type': {'YES': 'equipment', 'NO': 'spare'}.get(is_equip),
        'tag': text(H_TAG),
        'sap_material_number': text(H_SAP_MATNR),
        'material_category': text(H_CAT),
        'new_description': text(H_NEW_DESC),
        'maintenance_plant': text(H_PLANT),
        'manufacturer_name': text(H_MFR_NAME),
        'manufacturer_country_name': text(H_COUNTRY_NAME),
    }


def backfill_part_output_rows():
    """One-off catch-up for jobs made before part_output_rows existed:
    reads each such job's saved OUTPUT file and records its rows. A missing
    or unreadable file is skipped (and retried on the next startup).

    Columns are found by header name, so files from earlier versions of
    this layout still load: those name the sheet 'Sheet1' (read as the
    first sheet) and may lack 'Tag Number' (left blank) or 'NEW
    DESCRIPTION OF PARTS' (their 'Material Description' is used instead)."""
    for job_id, output_file in db.jobs_missing_part_output_rows():
        path = os.path.join(db.job_dir(job_id), output_file)
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except (OSError, KeyError, ValueError):
            continue
        try:
            ws = wb['BOM_WORKING'] if 'BOM_WORKING' in wb.sheetnames else wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
        finally:
            wb.close()
        col = {h: i for i, h in enumerate(rows[0]) if h is not None} if rows else {}
        if H_NEW_DESC not in col and H_MAT_DESC in col:
            col[H_NEW_DESC] = col[H_MAT_DESC]
        out = []
        for values in rows[5:]:   # 5 header rows, data from row 6
            r = _part_output_row(lambda h: values[col[h]] if h in col and col[h] < len(values) else None)
            if isinstance(r['material_temp_number'], (int, float)):
                out.append(r)
        db.record_part_output_rows(job_id, out)


def build_sap_output(parsed: dict, out_path: str, spir_filename: str = None, job_id: str = None):
    """spir_filename: the uploaded SPIR file's base name (no extension), used
    for CDC_SPIR NUMBER per the existing filename convention (app/main.py's
    `base = os.path.splitext(file.filename)[0]`). Falls back to the SPIR
    number parsed from inside the file if not given (e.g. direct/test use).

    job_id: this run's job id, used as the key for the central Part Number
    -> Material Temp Number registry (see engine.part_master) -- every
    equipment Model Number / spare Manufacturer Part Number resolved while
    building this file is permanently recorded against it. Required for
    real runs; falls back to spir_filename or spir_no for direct/test use
    where no job has been created."""
    cdc_spir_number = spir_filename or parsed['spir_no']
    series_job_id = job_id or spir_filename or parsed['spir_no'] or 'unknown'

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'BOM_WORKING'

    for c, h in enumerate(HEADERS, 1):
        ws.cell(row=1, column=c, value=h)
    for c, code in enumerate(SAP_CODES, 1):
        if code is not None:
            ws.cell(row=2, column=c, value=code)
    for c, ln in enumerate(FIELD_LENS, 1):
        if ln is not None:
            ws.cell(row=3, column=c, value=ln)
    for c, v in enumerate(ROW4, 1):
        if v is not None:
            ws.cell(row=4, column=c, value=v)
    for c in range(1, len(HEADERS) + 1):
        ws.cell(row=5, column=c, value=c)
    _apply_header_styling(ws, ncols=len(HEADERS))
    ws.freeze_panes = 'A6'   # freeze all 5 header rows

    series = PersistentMaterialNumberSeries(series_job_id)
    sheet_index = {sn: i + 1 for i, sn in enumerate(parsed['sheet_names'])}

    plant_code = reference_data.plant_code_for_spir_no(parsed['spir_no'])
    if not plant_code:
        _log_review(cdc_spir_number, '', parsed['manufacturer'], 'MAINTENANCE PLANNING PLANT',
                    f"no Reference Excel plant matches project number in SPIR NO '{parsed['spir_no']}'")

    rev = parsed.get('spir_rev')
    if not rev:
        _log_review(cdc_spir_number, '', parsed['manufacturer'], 'SPIR VS TAG REV',
                    "no 'ISSUE LETTER:' value found in the SPIR workbook")
    spir_type_code = spir_type_short_code(parsed['spir_type'])
    if not spir_type_code:
        _log_review(cdc_spir_number, '', parsed['manufacturer'], 'SPIR VS TAG SPIR TYPE',
                    'SPIR type not exactly one of the 4 known types (none/multiple selected)')
    generation_date = datetime.date.today()

    # Duplicate equipment (same Model Number + Manufacturer Part Number,
    # after Main Sheet + Continuation + Annexure resolution): only the
    # first TAG for a given combination gets a BOM_WORKING record; every
    # later TAG with the same combination is recorded in SPIR VS TAG
    # instead, referencing the first TAG's Material Temp Number as SUBMT.
    duplicate_of = _find_duplicate_tags(parsed)
    equip_id_by_tag = {}
    spir_vs_tag_rows = []

    row_out = 6
    serial = 0
    part_output_rows = []

    def write_row(row: dict):
        nonlocal row_out
        for c, h in enumerate(HEADERS, 1):
            ws.cell(row=row_out, column=c, value=row.get(h))
        row_out += 1
        part_output_rows.append(_part_output_row(row.get))

    for tag in parsed['tag_order']:
        info = parsed['tag_info'][tag]
        model = info['model']

        if tag in duplicate_of:
            submt = equip_id_by_tag.get(duplicate_of[tag])
            spir_vs_tag_rows.append([
                tag, submt, parsed['spir_no'], rev or None, parsed['manufacturer'], model,
                info['sern'], None, generation_date, None, spir_type_code or None, None, None
            ])
            continue

        equip_id = series.equipment_id(model, tag)
        equip_id_by_tag[tag] = equip_id

        # SPIR VS TAG is an all-TAG tracking sheet: every TAG (duplicate or
        # not) gets a row here, independently of BOM_WORKING's own
        # deduplication below. A non-duplicate TAG uses its own Equipment
        # Header Number as SUBMT (a duplicate TAG's row, built above,
        # already uses the first/owning TAG's number instead).
        spir_vs_tag_rows.append([
            tag, equip_id, parsed['spir_no'], rev or None, parsed['manufacturer'], model,
            info['sern'], None, generation_date, None, spir_type_code or None, None, None
        ])

        serial += 1
        equip_desc = equipment_new_description(parsed['equipment_desc'], model, parsed['manufacturer'])
        equip_desc_short = reference_data.abbreviate_description(equip_desc, 40)
        equip_country_name, equip_country_code = resolve_country(parsed['manufacturer'])
        if not equip_country_code:
            _log_review(cdc_spir_number, tag, parsed['manufacturer'], 'Manufacturer COUNTRY NAME / CODE',
                        'manufacturer country could not be resolved with confidence')

        equip_material_type = reference_data.material_type_code_for_text(
            f"{parsed['equipment_desc']} {model} {parsed['manufacturer']}")
        if not equip_material_type:
            _log_review(cdc_spir_number, tag, parsed['manufacturer'], 'MM REQUIREMENT MATERIAL TYPE',
                        'no unambiguous Reference Excel Material type code match for this equipment')

        # If one of this tag's own spare/item rows has the same Mfr Part
        # Number as the equipment's Model Number, that line IS the equipment
        # itself listed as a spare -- back-fill the equipment's otherwise
        # always-blank SAP Material Number / Delivery / Price / Currency
        # from it.
        self_item = equipment_self_item(parsed, tag, model)
        equip_sap_matnr = equip_delivery = equip_qar_price = equip_currency_code = None
        if self_item is not None:
            equip_sap_matnr = self_item['sap_no'] if self_item['sap_no'] not in (None, '') else None
            equip_delivery = self_item['delivery_wks']
            equip_qar_price, equip_currency_code = _qar_price(self_item['unit_price'], self_item['currency'])
            equip_currency_code = equip_currency_code or None
            if equip_qar_price is None and isinstance(self_item['unit_price'], (int, float)) and self_item['currency']:
                _log_review(cdc_spir_number, tag, parsed['manufacturer'], 'MM REQUIREMENT MOVING AVERAGE PRICE',
                            f"could not convert equipment's own currency to QAR for item {self_item['item_no']}")

        write_row({
            H_SNO: serial, H_TAG: tag, H_SAP_MATNR: equip_sap_matnr, H_MAT_TEMP: equip_id, H_CAT: 'B',
            H_OLD_MATNR: f'QEQU-{equip_id}', H_NEW_DESC: equip_desc, H_MAT_DESC: equip_desc_short,
            H_PLANT: plant_code or None, H_UOM: 'EA', H_MFR_PARTNO: model, H_MFR_NAME: parsed['manufacturer'],
            H_IS_EQUIP: 'YES', H_BOM_HDR_NUM: equip_id, H_QTY: 1, H_POS_NUM: '0010',
            H_CDC_SPIR: cdc_spir_number, H_DELIVERY: equip_delivery,
            H_COUNTRY_CODE: equip_country_code or None, H_MM_TYPE: equip_material_type or None,
            H_MM_GROUP: 'QEQU', H_MM_PRICE: equip_qar_price, H_COUNTRY_NAME: equip_country_name or None,
            H_CURRENCY: equip_currency_code,
        })

        pos_counter = 0
        for sn in parsed['sheet_names']:
            sheet = parsed['sheets'][sn]
            for it in sheet['items']:
                if tag not in it['flags']:
                    continue
                serial += 1
                pos_counter += 1
                pos_no = f'{pos_counter * 10:04d}'
                mat_id = series.spare_id(it['mfr_part_no'], sn, it['item_no'], tag)
                sap_matnr = it['sap_no'] if it['sap_no'] not in (None, '') else None
                qar_price, currency_code = _qar_price(it['unit_price'], it['currency'])
                if qar_price is None and isinstance(it['unit_price'], (int, float)) and it['currency']:
                    _log_review(cdc_spir_number, tag, parsed['manufacturer'], 'MM REQUIREMENT MOVING AVERAGE PRICE',
                                f"could not convert {currency_code!r} to QAR for item {it['item_no']}")
                new_desc = spare_new_description(it)
                new_desc_short = reference_data.abbreviate_description(new_desc, 40)
                spf_code = spf_number(sheet['spir_no'], sheet_index[sn], it['item_no'])
                spare_country_name, spare_country_code = resolve_country(it['supplier_ocm'])
                if not spare_country_code:
                    _log_review(cdc_spir_number, tag, it['supplier_ocm'], 'Manufacturer COUNTRY NAME / CODE',
                                f"manufacturer country could not be resolved with confidence for item {it['item_no']}")

                material_cat = reference_data.category_for_sap_no(it['sap_no'])
                material_group = reference_data.material_group_for_classification(
                    it['classification'], default_code='QPSP')

                # BOM Quantity is the actual per-tag quantity from this
                # item's own flag cell (C:F) for this tag -- NOT
                # 'qty_fitted' (column H, "TOTAL NO. OF IDENTICAL PARTS
                # FITTED"), which is a different, equipment-level count.
                # Flagged-but-non-numeric is a genuine "can't confidently
                # identify" case -- left blank and logged, never guessed.
                bom_qty = it['tag_qty'].get(tag)
                if bom_qty is None:
                    _log_review(cdc_spir_number, tag, parsed['manufacturer'], 'Quantity',
                                f"item {it['item_no']}'s tag-flag cell has no numeric quantity to use as BOM Quantity")

                write_row({
                    H_SNO: serial, H_TAG: tag, H_SAP_MATNR: sap_matnr, H_MAT_TEMP: mat_id, H_CAT: material_cat,
                    H_OLD_MATNR: spf_code, H_NEW_DESC: new_desc, H_MAT_DESC: new_desc_short,
                    H_PLANT: plant_code or None, H_UOM: str(it['uom'] or '').split(' ')[0],
                    H_MFR_PARTNO: it['mfr_part_no'], H_MFR_NAME: it['supplier_ocm'], H_IS_EQUIP: 'NO',
                    H_BOM_HDR_NUM: equip_id, H_QTY: bom_qty, H_POS_NUM: pos_no,
                    H_CDC_SPIR: cdc_spir_number, H_DELIVERY: it['delivery_wks'],
                    H_COUNTRY_CODE: spare_country_code or None, H_MM_TYPE: equip_material_type or None,
                    H_MM_GROUP: material_group, H_MM_PRICE: qar_price,
                    H_COUNTRY_NAME: spare_country_name or None,
                    H_CURRENCY: str(it['currency'] or '').split(' ')[0],
                })

    def col(header):
        return HEADERS.index(header) + 1

    last_row = row_out - 1
    for rr in range(6, last_row + 1):
        f_val = ws.cell(row=rr, column=col(H_OLD_MATNR)).value
        if f_val:
            ws.cell(row=rr, column=col(H_LEN_OLD), value=len(str(f_val)))
        i_val = ws.cell(row=rr, column=col(H_MAT_DESC)).value
        if i_val:
            ws.cell(row=rr, column=col(H_LEN_DESC), value=len(str(i_val)))
        m_val = ws.cell(row=rr, column=col(H_MFR_PARTNO)).value
        if m_val:
            ws.cell(row=rr, column=col(H_LEN_MFR), value=len(str(m_val)))
        o_val = ws.cell(row=rr, column=col(H_MFR_NAME)).value
        if o_val:
            ws.cell(row=rr, column=col(H_LEN_MFR_NAME), value=len(str(o_val)))

    # column widths, reasonable defaults
    for i in range(1, len(HEADERS) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 16

    if last_row >= 6:
        # A workbook whose every tag column failed to resolve to any real
        # tag (e.g. a malformed/unsupported reference) leaves BOM_WORKING
        # with zero data rows -- last_row < 6 then, which openpyxl's
        # range parser rejects outright ('D6:D5' is invalid). Guard here
        # so that always surfaces as the real "no tags found" problem
        # (see parse_spir's own empty-tag_order check) rather than a
        # confusing crash inside conditional formatting.
        dxf = DifferentialStyle(font=DUP_FONT, fill=DUP_FILL)
        mat_temp_col = openpyxl.utils.get_column_letter(col(H_MAT_TEMP))
        mfr_partno_col = openpyxl.utils.get_column_letter(col(H_MFR_PARTNO))
        ws.conditional_formatting.add(f'{mat_temp_col}6:{mat_temp_col}{last_row}', Rule(type='duplicateValues', dxf=dxf))
        ws.conditional_formatting.add(f'{mfr_partno_col}6:{mfr_partno_col}{last_row}', Rule(type='duplicateValues', dxf=dxf))

    ws['A5'].comment = Comment(
        "Material Temp Numbers are generated placeholders (NOT real SAP material numbers), "
        "permanently assigned per Part Number in a central registry shared across every SPIR: "
        "Equipment (Type B) starts at 40001+, reused whenever the same Model Number repeats -- "
        "in this SPIR or any other. Spares (Type L) start at 500001+, reused whenever the same "
        "Manufacturer's Part Number repeats -- in this SPIR or any other. "
        "QatarEnergy's SAP master-data team must assign the real numbers before upload.",
        "BOM Tool")

    _build_spir_vs_tag_sheet(wb, spir_vs_tag_rows)

    wb.save(out_path)
    db.record_part_output_rows(series_job_id, part_output_rows)
    return out_path
