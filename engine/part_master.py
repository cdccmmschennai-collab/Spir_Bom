"""
Central, persistent Part Number -> Material Temp Number registry (see
db.py for the underlying tables and the atomic get-or-create logic).

A Part Number's Material Temp Number is assigned once -- the first time
that part number is ever seen anywhere in the system -- and reused forever
after, for both equipment (keyed by Model Number) and spares (keyed by
Manufacturer's Part Number). This replaces the old MaterialNumberSeries in
rules.py, which only tracked numbers in memory for the duration of a single
upload.
"""
from .rules import normalize_part_value
from . import db


class PersistentMaterialNumberSeries:
    """One instance per job/upload. Resolves each equipment/spare Part
    Number against the central database (creating a new Material Temp
    Number only the first time that part number is seen anywhere) and
    records this job as one of the SPIRs that used it."""

    def __init__(self, job_id: str):
        self.job_id = job_id
        self._equip_cache = {}
        self._spare_cache = {}

    def equipment_id(self, model, tag: str) -> int:
        """Keyed by Model Number -- the same equipment model always gets
        the same Material Temp Number, whichever tag or SPIR it appears
        under. A tag with no usable Model Number falls back to a key
        scoped to that tag alone, so it still gets a stable number without
        being wrongly merged with every other blank-model tag."""
        key = normalize_part_value(model) or f'__TAG__:{tag}'
        if key in self._equip_cache:
            return self._equip_cache[key]
        mtn, _is_new, master_id = db.get_or_create_material_temp_number(
            key, 'equipment', display_value=str(model or tag).strip())
        db.record_spir_part_number(self.job_id, master_id, 'equipment', tag, mtn)
        self._equip_cache[key] = mtn
        return mtn

    def spare_id(self, mfr_part_no, sheet_name: str, item_no, tag: str) -> int:
        """Keyed by Manufacturer's Part Number -- the same part always
        gets the same Material Temp Number, whichever SPIR/tag it appears
        under. A spare with no usable Mfr Part Number falls back to a key
        scoped to that exact item, so blank-part-number spares are never
        wrongly merged into one shared number."""
        key = normalize_part_value(mfr_part_no) or f'__ITEM__:{self.job_id}:{sheet_name}:{item_no}:{tag}'
        if key in self._spare_cache:
            return self._spare_cache[key]
        mtn, _is_new, master_id = db.get_or_create_material_temp_number(
            key, 'spare', display_value=str(mfr_part_no or '').strip())
        db.record_spir_part_number(self.job_id, master_id, 'spare', tag, mtn)
        self._spare_cache[key] = mtn
        return mtn
