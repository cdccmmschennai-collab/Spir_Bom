"""
The rules we iterated on across the whole conversation, centralized so both
the Extraction writer and the SAP OUTPUT writer apply them identically.
"""


def expand_tag(tag: str) -> list:
    """'PM-3427A/B' -> ['PM-3427A', 'PM-3427B']
    A tag column can name several physical tags at once, sharing one common
    prefix with only the trailing letter(s) changing after each '/'. Tags
    with no '/' are returned unchanged as a single-item list."""
    tag = str(tag).strip()
    if '/' not in tag:
        return [tag]
    parts = [p.strip() for p in tag.split('/') if p.strip() != '']
    if not parts:
        return [tag]
    base = parts[0]
    result = [base]
    for suffix in parts[1:]:
        prefix = base[:-len(suffix)] if len(suffix) < len(base) else ''
        result.append(prefix + suffix)
    return result


def spir_numeric_segments(spir_no: str) -> list:
    """'VEN-4391-MEWTP-5-43-2003' -> ['4391', '5', '43', '2003']
    Keeps only the SPIR number's numeric '-'-separated segments, in order,
    dropping text segments (the leading 'VEN', a facility code like
    'MEWTP', etc.) -- dynamic for any SPIR number, not tied to a specific
    project or facility code."""
    return [p for p in str(spir_no or '').split('-') if p.strip().isdigit()]


def _join_segments_to_fit(segments: list, suffix: str, max_len: int) -> str:
    """Joins `segments` with '-' and appends `suffix`. If that exceeds
    max_len, the SPIR number's own middle segments are merged together
    (their separating hyphen dropped) one pair at a time, left to right,
    starting right after the leading segment, until it fits -- e.g.
    [4391,5,43,2003] -> [4391,543,2003] -> [4391,5432003]. The leading
    segment and `suffix` are never touched or truncated. Only if every
    internal hyphen has already been merged away and it still doesn't fit
    is the merged remainder trimmed from its front, as an absolute last
    resort to respect the hard max_len (the Excel field width)."""
    segs = list(segments)

    def build(s):
        return '-'.join(s) + suffix

    code = build(segs)
    while len(code) > max_len and len(segs) > 2:
        segs = [segs[0], segs[1] + segs[2]] + segs[3:]
        code = build(segs)

    if len(code) > max_len and len(segs) > 1:
        overflow = len(code) - max_len
        rest = segs[1][overflow:] if overflow < len(segs[1]) else ''
        segs = [segs[0], rest] if rest else [segs[0]]
        code = build(segs)

    # Absolute failsafe: only reachable with an unrealistically long leading
    # segment or suffix (real SPIR project numbers / line identifiers are
    # nowhere near this), but the Excel field's hard width limit must never
    # be violated regardless.
    if len(code) > max_len:
        code = code[:max_len]

    return code


def spf_prefix(spir_no: str, max_len: int = 18) -> str:
    """OLD MATERIAL NUMBER/SPF NUMBER for an equipment/header row: the SPIR
    number's own numeric segments (see spir_numeric_segments), compressed
    to fit max_len if needed."""
    segments = spir_numeric_segments(spir_no)
    if not segments:
        return str(spir_no or '').strip()[:max_len]
    return _join_segments_to_fit(segments, '', max_len)


def spf_number(spir_no: str, sheet_index: int, line_no, max_len: int = 18) -> str:
    """OLD MATERIAL NUMBER/SPF NUMBER for a spare row:
    '<SPIR number's numeric segments>-<sheet index>L<line, 2-digit>'
    e.g. 'VEN-4391-MEWTP-5-43-2003' + sheet 1, line 1
    -> '4391-5-43-2003-1L01' (19 chars, over the 18-char SAP limit)
    -> '4391-543-2003-1L01' (18 chars, via _join_segments_to_fit)."""
    segments = spir_numeric_segments(spir_no)
    suffix = f'-{sheet_index}L{int(line_no):02d}'
    if not segments:
        return (str(spir_no or '').strip() + suffix)[:max_len]
    return _join_segments_to_fit(segments, suffix, max_len)


def spare_new_description(item: dict) -> str:
    """'description,mfr part number,supplier' -- matches the Extraction file's
    'NEW DESCRIPTION OF PARTS' convention for spare (leaf) rows."""
    return f"{item['desc']},{item['mfr_part_no']},{item['supplier_ocm']}"


def equipment_new_description(equipment_desc: str, model, manufacturer: str) -> str:
    """'equipment name, model, equipment make' -- the equivalent convention
    for equipment/BOM-header rows."""
    return f"{equipment_desc}, {model}, {manufacturer}"


def normalize_part_value(v) -> str:
    """Trims/upper-cases a Model Number or Manufacturer Part Number for
    comparison only (never for display) -- '' for missing/blank/'-'."""
    s = str(v).strip() if v is not None else ''
    return '' if s in ('', '-') else s.upper()


def equipment_self_item(parsed, tag, model):
    """The spare/item row (if any) flagged for this tag whose Manufacturer
    Part Number equals the equipment's own Model Number -- i.e. the SPIR
    lists the equipment itself as one of its own spare/replacement lines.
    Used by both the Output and Extraction writers to back-fill the
    equipment row's own SAP Material Number / Delivery / Price / Currency
    (otherwise always blank, since the SPIR has no equipment-level source
    for them)."""
    model_key = normalize_part_value(model)
    if not model_key:
        return None
    for sn in parsed['sheet_names']:
        for it in parsed['sheets'][sn]['items']:
            if tag in it['flags'] and normalize_part_value(it['mfr_part_no']) == model_key:
                return it
    return None


def equipment_dedup_key(model, mfr_part_no):
    """Normalized (Model Number, Manufacturer Part Number) key used to detect
    duplicate equipment records. Returns None (never a duplicate) if either
    value is missing/blank/'-' -- a missing value must not be treated as a
    match."""
    m = normalize_part_value(model)
    p = normalize_part_value(mfr_part_no)
    if not m or not p:
        return None
    return (m, p)


SPIR_TYPE_SHORT_CODES = {
    'COMMISSIONING SPARES': 'CS',
    'INITIAL SPARES': 'IS',
    'NORMAL OPERATING SPARES': 'NOS',
    'LIFE CYCLE SPARES': 'LCS',
}


def spir_type_short_code(spir_type: str) -> str:
    """'NORMAL OPERATING SPARES' -> 'NOS'. Returns '' for anything not
    exactly one of the four known types (including blank/ambiguous) --
    never guessed."""
    return SPIR_TYPE_SHORT_CODES.get(str(spir_type or '').strip().upper(), '')


class MaterialNumberSeries:
    """Assigns Material Temp Numbers:
       - Equipment (Type B): one NEW number per physical tag, starting at 40001,
         even when the same model repeats across tags.
       - Spares (Type L): deduped by Manufacturer's Part Number -- the SAME
         number is reused every time an identical part repeats, starting at
         500001.
    """

    def __init__(self, equip_start: int = 40000, spare_start: int = 500000):
        self._equip_counter = equip_start
        self._spare_counter = spare_start
        self._equip_ids = {}   # tag -> id
        self._spare_ids = {}   # mfr_part_no -> id

    def equipment_id(self, tag: str) -> int:
        if tag not in self._equip_ids:
            self._equip_counter += 1
            self._equip_ids[tag] = self._equip_counter
        return self._equip_ids[tag]

    def spare_id(self, mfr_part_no) -> int:
        key = str(mfr_part_no).strip()
        if key not in self._spare_ids:
            self._spare_counter += 1
            self._spare_ids[key] = self._spare_counter
        return self._spare_ids[key]
