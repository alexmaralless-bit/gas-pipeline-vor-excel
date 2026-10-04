"""spec_writer.py — generated «Спец» sheet as a projection over «ВР».

The manual reference workbook does not keep «Спец» as an independent source of
truth: most rows are formulas pointing to «ВР» plus VLOOKUPs into «Каталог».
This module preserves that direction. The expanded «ВР» is the primary model;
«Спец» is derived from it and can later be rendered to PDF.
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .catalog_loader import CatalogBundle, CatalogRow


SHEET_SPEC = "Спец"
SHEET_VR = "ВР"
SHEET_CATALOG = "Каталог"


def _is_spec_candidate(name: str, note: str) -> bool:
    if not name:
        return False
    if "DB lookup" in note or "catalog:" in note:
        return True
    return name.startswith("Труба DN")


def _catalog_id_from_note(note: str) -> int | None:
    marker = "catalog_id="
    if marker not in note:
        return None
    tail = note.split(marker, 1)[1]
    raw = tail.split(",", 1)[0].split(" ", 1)[0].strip()
    try:
        return int(raw)
    except ValueError:
        return None


def _catalog_by_id(catalog: CatalogBundle | None) -> dict[int, CatalogRow]:
    if catalog is None:
        return {}
    return {row.row_id: row for row in catalog.rows}


def build_spec_sheet_from_vr(
    wb: Workbook,
    *,
    vr_sheet_name: str = SHEET_VR,
    catalog: CatalogBundle | None = None,
) -> None:
    """Create «Спец» sheet from generated «ВР» rows.

    Current scope is deliberately conservative:
      * include rows that are catalog-backed, plus pipe material rows;
      * reference ВР quantities by formula, so edits/recalculation flow through;
      * use Каталог VLOOKUPs for type/manufacturer/mass where possible.

    Later passes can add manual-style grouping of repeated catalog positions.
    """
    if vr_sheet_name not in wb.sheetnames:
        return
    if SHEET_SPEC in wb.sheetnames:
        del wb[SHEET_SPEC]

    vr = wb[vr_sheet_name]
    ws = wb.create_sheet(SHEET_SPEC, 0)
    catalog_rows = _catalog_by_id(catalog)

    headers = [
        "Позиция",
        "Наименование",
        "Тип, марка, обозначение документа, опросного листа",
        "Код оборудования, изделия, материала",
        "Завод-изготовитель",
        "Единица измере-ния",
        "Количество",
        "Масса единицы, кг",
        "Примечание",
        "Количество",
    ]
    for idx, header in enumerate(headers, start=3):
        cell = ws.cell(1, idx)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        ws.cell(2, idx).value = idx - 2

    out_row = 3
    ws.cell(out_row, 3).value = 1
    ws.cell(out_row, 4).value = "Оборудование и материалы"
    ws.cell(out_row, 4).font = Font(bold=True)
    out_row += 1

    pos = 2
    for vr_row in range(1, vr.max_row + 1):
        item_num = vr.cell(vr_row, 2).value
        name = str(vr.cell(vr_row, 4).value or "")
        note = str(vr.cell(vr_row, 9).value or "")
        unit = vr.cell(vr_row, 5).value
        qty = vr.cell(vr_row, 6).value
        if item_num in (None, "") or unit in (None, "") or qty in (None, ""):
            continue
        if not _is_spec_candidate(name, note):
            continue
        catalog_row = catalog_rows.get(_catalog_id_from_note(note) or -1)

        ws.cell(out_row, 3).value = pos
        ws.cell(out_row, 4).value = f"='{vr_sheet_name}'!D{vr_row}"
        ws.cell(out_row, 5).value = catalog_row.type_mark if catalog_row else None
        ws.cell(out_row, 6).value = catalog_row.code if catalog_row else None
        ws.cell(out_row, 7).value = catalog_row.manufacturer if catalog_row else None
        ws.cell(out_row, 8).value = f"='{vr_sheet_name}'!E{vr_row}"
        ws.cell(out_row, 9).value = f"='{vr_sheet_name}'!F{vr_row}"
        ws.cell(out_row, 10).value = catalog_row.mass_per_unit if catalog_row else None
        ws.cell(out_row, 11).value = f"='{vr_sheet_name}'!I{vr_row}"
        ws.cell(out_row, 12).value = f"=I{out_row}"
        out_row += 1
        pos += 1

    for col in range(3, 13):
        ws.column_dimensions[get_column_letter(col)].width = {
            3: 10, 4: 64, 5: 32, 6: 18, 7: 24, 8: 14, 9: 14, 10: 14, 11: 32, 12: 14
        }.get(col, 14)

    ws.freeze_panes = "C3"
