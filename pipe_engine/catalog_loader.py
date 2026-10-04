"""
catalog_loader.py — загрузка pipe_catalog.db в Python-объекты + helpers.

Используется vr_writer/spec_writer чтобы:
 * писать лист «Каталог» (995 строк × 13 колонок) в xlsx
 * писать лист «Тег» (pivot tag → min/max row)
 * строить VLOOKUP формулы для ВР components с catalog_match

Каталог — SQLite база в ``pipe_engine/catalog/pipe_catalog.db``,
извлечённая разовым скриптом из эталонной книги.
"""
from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


DEFAULT_DB_PATH = Path(__file__).resolve().parent.parent / "catalog" / "pipe_catalog.db"


# Каталог column order in xlsx (как у эталона — для совместимости VLOOKUP индексов):
#   1: DN
#   2: DN2
#   3: Поз (Позиция)
#   4: Наименование и техническая характеристика  ← target most VLOOKUPs
#   5: Тип, марка, обозначение
#   6: Код оборудования
#   7: Завод-изготовитель
#   8: Единица измерения
#   9: Количество
#   10: Масса единицы, кг
#   11: Примечание
#   12: Тэги
#   13: Поз (= ROW — формула в эталоне, у нас literal row number)

COL_NAME_IDX     = 4   # "Наименование" — самый частый target
COL_TYPE_IDX     = 5   # "Тип/марка"
COL_CODE_IDX     = 6   # "Код"
COL_VENDOR_IDX   = 7   # "Завод-изготовитель"
COL_UNIT_IDX     = 8
COL_MASS_IDX     = 10


@dataclass
class CatalogRow:
    row_id: int           # = catalog.id = original Excel row number
    dn: Optional[str]
    dn2: Optional[str]
    pos: Optional[str]
    name: Optional[str]
    type_mark: Optional[str]
    code: Optional[str]
    manufacturer: Optional[str]
    unit: Optional[str]
    qty: Optional[str]
    mass_per_unit: Optional[float]
    note: Optional[str]
    tag: Optional[str]


@dataclass
class TagPivotRow:
    tag: str
    min_pos: int        # = min catalog.id where tag appears
    max_pos: int        # = max catalog.id
    manufacturer_filter: Optional[str]


@dataclass
class DnTranslationRow:
    diameter_outer: int
    dn: int


@dataclass
class ClimateSp131Row:
    source_row: int
    region: Optional[str]
    locality: Optional[str]
    cold_day_098_c: Optional[float]
    cold_day_092_c: Optional[float]
    cold_5day_098_c: Optional[float]
    cold_5day_092_c: Optional[float]
    temp_094_c: Optional[float]
    abs_min_c: Optional[float]
    cold_month_daily_amplitude_c: Optional[float]
    duration_below_0_days: Optional[int]
    avg_temp_below_0_c: Optional[float]
    duration_below_8_days: Optional[int]
    avg_temp_below_8_c: Optional[float]
    duration_below_10_days: Optional[int]
    avg_temp_below_10_c: Optional[float]
    avg_monthly_relative_humidity_pct: Optional[float]
    avg_monthly_relative_humidity_15h_pct: Optional[float]
    precipitation_nov_mar_mm: Optional[float]
    dominant_wind_dec_feb: Optional[str]
    max_avg_wind_speed_jan_m_s: Optional[float]
    avg_wind_speed_below_8_m_s: Optional[float]


@dataclass
class WorkCatalogRow:
    work_key: str
    section: Optional[str]
    name_for_vor: str
    unit: str
    cvor_type: str
    formula_key: Optional[str]
    estimate_norm_code: Optional[str]
    estimate_norm_name: Optional[str]
    norm_source: Optional[str]
    source: Optional[str]
    note: Optional[str]
    active: bool


@dataclass
class CatalogBundle:
    """Everything we need from pipe_catalog.db."""
    rows: list[CatalogRow]
    tags: list[TagPivotRow]
    dn_translations: list[DnTranslationRow]
    climate_sp131: list[ClimateSp131Row]
    work_catalog: list[WorkCatalogRow]
    project_params: dict[str, str]   # key → value
    source_xlsm: Optional[str] = None

    def tag_range(self, tag_substring: str) -> Optional[tuple[int, int]]:
        """Find (min_pos, max_pos) for a tag matched by substring (case-insensitive).

        Returns the FIRST matching tag's range. None if no match.
        """
        needle = tag_substring.lower()
        for t in self.tags:
            if needle in t.tag.lower():
                return (t.min_pos, t.max_pos)
        return None

    def find_row_by_tag_and_dn(self, tag_substring: str,
                                dn: int | str) -> Optional[CatalogRow]:
        """Find first catalog row matching tag substring + DN (DN as string for compare)."""
        dn_str = str(dn)
        needle = tag_substring.lower()
        for r in self.rows:
            if r.tag and needle in r.tag.lower() and r.dn == dn_str:
                return r
        return None

    def work_by_key(self, work_key: str) -> Optional[WorkCatalogRow]:
        """Find an active work row by stable key."""
        for row in self.work_catalog:
            if row.active and row.work_key == work_key:
                return row
        return None


def load_catalog(db_path: str | Path = DEFAULT_DB_PATH) -> CatalogBundle:
    """Load full catalog from SQLite into in-memory CatalogBundle."""
    db_path = Path(db_path)
    if not db_path.exists():
        raise FileNotFoundError(f"catalog DB not found: {db_path}")

    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    try:
        rows = []
        for r in conn.execute("SELECT * FROM catalog ORDER BY id"):
            rows.append(CatalogRow(
                row_id=r["id"], dn=r["dn"], dn2=r["dn2"], pos=r["pos"],
                name=r["name"], type_mark=r["type_mark"], code=r["code"],
                manufacturer=r["manufacturer"], unit=r["unit"], qty=r["qty"],
                mass_per_unit=r["mass_per_unit"], note=r["note"], tag=r["tag"],
            ))
        tags = []
        for r in conn.execute("SELECT * FROM tag_pivot ORDER BY min_pos"):
            tags.append(TagPivotRow(
                tag=r["tag"], min_pos=r["min_pos"], max_pos=r["max_pos"],
                manufacturer_filter=r["manufacturer_filter"],
            ))
        dn_translations = []
        for r in conn.execute("SELECT * FROM dn_translation ORDER BY dn"):
            dn_translations.append(DnTranslationRow(
                diameter_outer=int(r["diameter_outer"]),
                dn=int(r["dn"]),
            ))
        climate_sp131 = []
        climate_exists = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='climate_sp131'"
        ).fetchone()
        if climate_exists:
            for r in conn.execute("SELECT * FROM climate_sp131 ORDER BY source_row"):
                climate_sp131.append(ClimateSp131Row(
                    source_row=int(r["source_row"]),
                    region=r["region"],
                    locality=r["locality"],
                    cold_day_098_c=r["cold_day_098_c"],
                    cold_day_092_c=r["cold_day_092_c"],
                    cold_5day_098_c=r["cold_5day_098_c"],
                    cold_5day_092_c=r["cold_5day_092_c"],
                    temp_094_c=r["temp_094_c"],
                    abs_min_c=r["abs_min_c"],
                    cold_month_daily_amplitude_c=r["cold_month_daily_amplitude_c"],
                    duration_below_0_days=r["duration_below_0_days"],
                    avg_temp_below_0_c=r["avg_temp_below_0_c"],
                    duration_below_8_days=r["duration_below_8_days"],
                    avg_temp_below_8_c=r["avg_temp_below_8_c"],
                    duration_below_10_days=r["duration_below_10_days"],
                    avg_temp_below_10_c=r["avg_temp_below_10_c"],
                    avg_monthly_relative_humidity_pct=r["avg_monthly_relative_humidity_pct"],
                    avg_monthly_relative_humidity_15h_pct=r["avg_monthly_relative_humidity_15h_pct"],
                    precipitation_nov_mar_mm=r["precipitation_nov_mar_mm"],
                    dominant_wind_dec_feb=r["dominant_wind_dec_feb"],
                    max_avg_wind_speed_jan_m_s=r["max_avg_wind_speed_jan_m_s"],
                    avg_wind_speed_below_8_m_s=r["avg_wind_speed_below_8_m_s"],
                ))
        work_catalog = []
        work_exists = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='work_catalog'"
        ).fetchone()
        if work_exists:
            for r in conn.execute("SELECT * FROM work_catalog ORDER BY work_key"):
                work_catalog.append(WorkCatalogRow(
                    work_key=r["work_key"],
                    section=r["section"],
                    name_for_vor=r["name_for_vor"],
                    unit=r["unit"],
                    cvor_type=r["cvor_type"],
                    formula_key=r["formula_key"],
                    estimate_norm_code=r["estimate_norm_code"],
                    estimate_norm_name=r["estimate_norm_name"],
                    norm_source=r["norm_source"],
                    source=r["source"],
                    note=r["note"],
                    active=bool(r["active"]),
                ))
        project_params = {}
        for r in conn.execute("SELECT key, value FROM project_params"):
            project_params[r["key"]] = r["value"]
        src = None
        for r in conn.execute("SELECT value FROM _meta WHERE key='source_xlsm'"):
            src = r["value"]
            break
        return CatalogBundle(rows=rows, tags=tags,
                             dn_translations=dn_translations,
                             climate_sp131=climate_sp131,
                             work_catalog=work_catalog,
                             project_params=project_params, source_xlsm=src)
    finally:
        conn.close()
