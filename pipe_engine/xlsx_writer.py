"""
xlsx_writer.py — emit an **earthwork-calculation xlsx** where every
computed cell is an Excel formula string, not a pre-computed number.

NOTE: what this module produces is the «Расчет земли» part of a full
ВОР workflow (mirroring the corresponding sheet of the reference workbook).
Other sheets (Спец, ВР, Тег) and the actual ВОР ledger live higher up
and require catalog-lookup integration — out of scope for step 4.

Non-negotiable design constraint (user requirement): the engineer must
be able to open the xlsx, see the formulas, audit them, modify any
input cell, and have Excel recalculate everything automatically. The
Python side computes the same numbers in ``pipe_engine.earthwork``
only for regression testing — the xlsx itself contains formulas as
cell values.

Layout (two sheets — Спец/ВР sheets are out of scope for this module):

* **«ИсхДанные»** — 8-row input panel. Engineer edits B-column to
  re-parameterize the project: depth, sand pad, snow, trench bottom
  width, slope ratio, sand-cushion flag, season flag.
* **«Расчет земли»** — main table. One row per (branch, DN). Per-row
  formulas for I/J/K/L/M reference ИсхДанные. ИТОГО row sums them.
  Earthworks summary block below references the ИТОГО.

Inputs come from ``pipe_engine.trench_profile`` and ``pipe_engine.earthwork``
dataclasses; the function does NOT re-run the math — it just writes
formulas with cell-references.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .catalog_loader import CatalogBundle, load_catalog
from .cvor_writer import build_cvor_sheets
from .earthwork import ExitSandNode, PitParams, ProjectParams
from .io_schema import BlockRecord
from .node_library import NodeLibrary
from .spec_writer import build_spec_sheet_from_vr
from .trench_profile import Branch, JointPipePair
from .vr_writer import build_vr_sheet


SHEET_CATALOG = "Каталог"
SHEET_TAG     = "Тег"


# ============================================================================
# ИсхДанные sheet layout — single source of truth for cell positions
# ============================================================================

# Each constant is the ROW number on the ИсхДанные sheet (column B).
# Keep these in sync with build_inputs_sheet() below.
IDX_DEPTH               = 1   # B1 — глубина заложения, м
IDX_SAND_PAD            = 2   # B2 — толщина песчаной подушки/присыпки, м
IDX_SNOW_THICKNESS      = 3   # B3 — толщина снежного покрова, м
IDX_BOTTOM_WIDTH        = 4   # B4 — ширина траншеи понизу (одиночная), м
IDX_SLOPE_RATIO         = 5   # B5 — уклон откоса (1:k)
IDX_SAND_CUSHION_FLAG   = 6   # B6 — 1=песчаная подушка нужна, 0=нет
IDX_SEASON_WINTER_FLAG  = 7   # B7 — 1=зима (расчёт снега), 0=лето
IDX_HAND_TRENCH_COUNT   = 8   # B8 — кол-во ручных траншей под пересечения
IDX_HAND_TRENCH_LEN     = 9   # B9 — длина одного участка ручной траншеи, м
IDX_HAND_TRENCH_WIDTH   = 10  # B10 — ширина ручной траншеи, м (обычно = bottom_width)
IDX_JOINT_BOTTOM_WIDTH  = 11  # B11 — ширина траншеи понизу для совместной прокладки, м
IDX_SAND_PER_GNB_EXIT   = 12  # B12 — формула 0.35*0.35*B1
IDX_PIPE_OVERRUN_FACTOR = 13  # B13 — overrun трубы (змейка + обрезки), default 1.10

SHEET_INPUTS = "ИсхДанные"
SHEET_TRACE  = "Расчет земли"


def _inp(addr_row: int) -> str:
    """Build an absolute cell reference to the ИсхДанные sheet, e.g. 'ИсхДанные!$B$1'."""
    return f"{SHEET_INPUTS}!$B${addr_row}"


# ============================================================================
# Inputs sheet
# ============================================================================

def build_inputs_sheet(wb: Workbook, params: ProjectParams,
                       bottom_width_default: float,
                       slope_ratio_default: float,
                       sand_cushion_required: bool,
                       season_winter: bool,
                       hand_trench_count: int = 0,
                       hand_trench_section_length: float = 4.0,
                       hand_trench_width: float = 0.7,
                       joint_bottom_width: float = 1.2,
                       pipe_overrun_factor: float = 1.10) -> None:
    """Create the ИсхДанные sheet with one input per row.

    All cells in column B can be edited by the engineer in Excel —
    formulas elsewhere reference these via absolute ИсхДанные!$B$N refs.
    """
    ws = wb.create_sheet(SHEET_INPUTS, 0)

    header_font = Font(bold=True)
    fill = PatternFill("solid", fgColor="FFE699")

    rows = [
        ("Глубина заложения, м",                          float(params.depth)),
        ("Подсыпка песка, м",                             float(params.sand_pad)),
        ("Толщина снежного покрова, м",                   float(params.snow_thickness)),
        ("Ширина траншеи понизу (одиночная), м",          float(bottom_width_default)),
        ("Уклон откоса (1:k, 0=вертикально)",             float(slope_ratio_default)),
        ("Песчаная подушка нужна (1=да, 0=нет)",          1 if sand_cushion_required else 0),
        ("Сезон (1=winter, 0=summer)",                    1 if season_winter else 0),
        ("Кол-во ручных траншей (пересечения)",           int(hand_trench_count)),
        ("Длина 1 ручной траншеи, м",                     float(hand_trench_section_length)),
        ("Ширина ручной траншеи, м",                      float(hand_trench_width)),
        ("Ширина траншеи понизу (совместная), м",         float(joint_bottom_width)),
        ("Объем песка на 1 выход, м³",                    f"=0.35*0.35*$B${IDX_DEPTH}"),
        ("Overrun трубы (змейка + обрезки)",              float(pipe_overrun_factor)),
    ]
    for r, (label, value) in enumerate(rows, start=1):
        ws.cell(r, 1).value = label
        ws.cell(r, 2).value = value
        ws.cell(r, 2).fill = fill
        if r == IDX_SAND_PER_GNB_EXIT:
            ws.cell(r, 2).fill = PatternFill("solid", fgColor="E2EFDA")  # derived — different colour

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 14
    for r in range(1, len(rows) + 1):
        ws.cell(r, 1).font = header_font


# ============================================================================
# Trace sheet — main table with per-row formulas
# ============================================================================

# Trace sheet column layout — keep in sync with build_trace_sheet().
COL_BRANCH_ID   = 1   # A — branch.id (trace text)
COL_DN          = 2   # B — DN (mm)
COL_TRACE_LEN   = 3   # C — total trace length (m)
COL_GNB_LEN     = 4   # D — GNB length (m)
COL_JOINT_LEN   = 5   # E — length in joint laying (SUMIF over joint block)
COL_OPEN_LEN    = 6   # F — open-trench length (= C - D - E)
COL_PIPE_VOL    = 7   # G — J: pipe internal volume (formula)
COL_SOIL_VOL    = 8   # H — K: trench soil volume (formula)
COL_SAND_SIDE   = 9   # I — L: sand at sides (formula)
COL_SAND_CUSHION= 10  # J — M: sand cushion + sprinkle base (formula)
COL_NOTE        = 11  # K — free-form note

TRACE_HEADER_ROW = 1
TRACE_FIRST_DATA_ROW = 2   # data starts immediately below the header


def _flat_pipe_rows(branches: Iterable[Branch]) -> list[tuple[str, int, float, float]]:
    """Flatten (branch_id, dn, total_length, gnb_length) for every PipeRow."""
    out = []
    for b in branches:
        for p in b.pipes:
            total = p.length_grp1 + p.length_other
            out.append((b.id, p.dn_mm, total, p.gnb_length))
    return out


def build_trace_sheet(wb: Workbook, branches: list[Branch],
                       joint_pair_count: int = 0) -> tuple[int, int]:
    """Build the «Расчет земли» sheet with per-(branch,DN) formula rows.

    ``joint_pair_count`` is the number of joint-pipe pairs that will be
    appended later via ``build_joint_pipe_block``; we use it to forward-
    compute the addresses for the joint-length SUMIF lookup in column E.
    Pass 0 if there are no joint pairs.

    Returns (first_data_row, итого_row) so callers can build downstream
    formulas (e.g. earthworks summary) that reference ИТОГО.
    """
    ws = wb.create_sheet(SHEET_TRACE)

    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    formula_fill = PatternFill("solid", fgColor="F2F2F2")
    итого_fill   = PatternFill("solid", fgColor="FFD966")

    # ---- header ---------------------------------------------------------
    headers = [
        ("A", "Ветка"),
        ("B", "Ду, мм"),
        ("C", "Длина трассы, м"),
        ("D", "Длина ГНБ, м"),
        ("E", "Совместная прокладка, м"),
        ("F", "Длина открытой траншеи I, м"),
        ("G", "Объём труб J, м³"),
        ("H", "Объём грунта K, м³"),
        ("I", "Песок по сторонам L, м³"),
        ("J", "Подушка + присыпка M, м³"),
        ("K", "Примечание"),
    ]
    for col_letter, text in headers:
        c = ws[f"{col_letter}{TRACE_HEADER_ROW}"]
        c.value = text
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[TRACE_HEADER_ROW].height = 36
    for col, width in [("A", 28), ("B", 8), ("C", 14), ("D", 12), ("E", 14),
                       ("F", 16), ("G", 14), ("H", 14), ("I", 14), ("J", 16), ("K", 24)]:
        ws.column_dimensions[col].width = width

    # ---- data rows ------------------------------------------------------
    rows = _flat_pipe_rows(branches)
    r = TRACE_FIRST_DATA_ROW
    first_data_row = r

    # Forward-compute joint-block address range (filled by build_joint_pipe_block).
    # That function lays out: trace_итого + 2 = section title, +3 = header,
    # +4 = first data row, then N data rows, then ИТОГО.
    n_trace_rows = len(rows)
    trace_итого_row = TRACE_FIRST_DATA_ROW + n_trace_rows
    joint_first_data_row = trace_итого_row + 4   # title=+2, header=+3, data starts=+4
    joint_last_data_row  = joint_first_data_row + max(0, joint_pair_count - 1)
    # In joint block: B=DN1, C=DN2, E=length (see build_joint_pipe_block)
    has_joints = joint_pair_count > 0

    for branch_id, dn_mm, total_len, gnb_len in rows:
        # Literal inputs (engineer can override)
        ws.cell(r, COL_BRANCH_ID).value = branch_id
        ws.cell(r, COL_DN).value = dn_mm
        ws.cell(r, COL_TRACE_LEN).value = round(total_len, 3)
        ws.cell(r, COL_GNB_LEN).value = round(gnb_len, 3)

        # E = joint length: SUMIF over the joint block by DN match (both columns B and C)
        if has_joints:
            jb = f"B{joint_first_data_row}:B{joint_last_data_row}"
            jc = f"C{joint_first_data_row}:C{joint_last_data_row}"
            jlen = f"E{joint_first_data_row}:E{joint_last_data_row}"
            ws.cell(r, COL_JOINT_LEN).value = (
                f"=SUMIF({jb},B{r},{jlen})+SUMIF({jc},B{r},{jlen})"
            )
        else:
            # Still a formula (not literal 0) so the schema-invariant test holds.
            ws.cell(r, COL_JOINT_LEN).value = "=0"

        # F = open length = C - D - E
        ws.cell(r, COL_OPEN_LEN).value = f"=C{r}-D{r}-E{r}"

        # G = pipe internal volume J = PI() * (DN/1000)^2 / 4 * I (open length)
        ws.cell(r, COL_PIPE_VOL).value = f"=PI()*(B{r}/1000)^2/4*F{r}"

        # ---- K (soil volume, single trench) — trapezoidal ----
        depth_total = (
            f"({_inp(IDX_DEPTH)}+B{r}/1000+IF({_inp(IDX_SAND_CUSHION_FLAG)}=1,{_inp(IDX_SAND_PAD)},0))"
        )
        area = f"({_inp(IDX_BOTTOM_WIDTH)}+{depth_total}*{_inp(IDX_SLOPE_RATIO)})*{depth_total}"
        ws.cell(r, COL_SOIL_VOL).value = f"=F{r}*{area}"

        # ---- L (sand at sides) — only if sand_cushion required ----
        ws.cell(r, COL_SAND_SIDE).value = (
            f"=IF({_inp(IDX_SAND_CUSHION_FLAG)}=1,"
            f"(B{r}/1000*{_inp(IDX_BOTTOM_WIDTH)}*C{r}-G{r}),0)"
        )

        # ---- M (sand cushion+sprinkle base) ----
        ws.cell(r, COL_SAND_CUSHION).value = (
            f"=IF({_inp(IDX_SAND_CUSHION_FLAG)}=1,"
            f"F{r}*{_inp(IDX_BOTTOM_WIDTH)}*{_inp(IDX_SAND_PAD)},0)"
        )

        for col in (COL_JOINT_LEN, COL_OPEN_LEN, COL_PIPE_VOL, COL_SOIL_VOL,
                    COL_SAND_SIDE, COL_SAND_CUSHION):
            ws.cell(r, col).fill = formula_fill
        r += 1

    last_data_row = r - 1
    итого_row = r

    # ---- ИТОГО ---------------------------------------------------------
    ws.cell(итого_row, COL_BRANCH_ID).value = "ИТОГО"
    ws.cell(итого_row, COL_BRANCH_ID).font = Font(bold=True)
    for col in (COL_TRACE_LEN, COL_GNB_LEN, COL_JOINT_LEN, COL_OPEN_LEN,
                COL_PIPE_VOL, COL_SOIL_VOL, COL_SAND_SIDE, COL_SAND_CUSHION):
        letter = get_column_letter(col)
        ws.cell(итого_row, col).value = f"=SUM({letter}{first_data_row}:{letter}{last_data_row})"
        ws.cell(итого_row, col).font = Font(bold=True)
        ws.cell(итого_row, col).fill = итого_fill

    return first_data_row, итого_row


# ============================================================================
# Earthworks summary block (appended below ИТОГО)
# ============================================================================

def append_earthworks_summary(
    wb: Workbook,
    итого_row: int,
    exit_sand_total_cell_ref: Optional[str],
    pits: Optional[PitParams],
    joint_итого_row: Optional[int] = None,
) -> dict:
    """Append rows below ИТОГО with snow / hand-trench / excavator / sand /
    berm / balance formulas referencing ИТОГО cells.

    For MVP, exit_sand and pit blocks are kept inline (not on separate
    sheets). ``exit_sand_total_cell_ref`` — if provided, used directly
    in the c62 sum formula; if None, we drop that term.

    Hand-trench count/length/width are now in ИсхДанные (rows 8-10) — the
    formula references them, not Python params.
    """
    ws = wb[SHEET_TRACE]

    # Start AFTER any previously appended block (e.g. exit_sand) — not just
    # below ИТОГО, because the caller may have written more rows already.
    start = max(итого_row + 2, ws.max_row + 2)
    r = start

    # Totals on the trace ИТОГО row.
    # After insertion of «Совместная прокладка» column (E), the layout is:
    # F=open_len(I) G=pipe_vol(J) H=soil_vol(K) I=sand_side(L) J=sand_cushion(M)
    I_TOTAL_TRACE = f"F{итого_row}"
    J_TOTAL_TRACE = f"G{итого_row}"
    K_TOTAL_TRACE = f"H{итого_row}"
    L_TOTAL_TRACE = f"I{итого_row}"
    M_TOTAL_TRACE = f"J{итого_row}"

    # If a joint-pipe block exists, sum trace + joint for each total via
    # ranges like "H{trace_итого}+H{joint_итого}". The joint block uses
    # the same column letters per build_joint_pipe_block layout.
    def _combine(tcell: str, col_letter: str) -> str:
        if joint_итого_row is not None:
            return f"({tcell}+{col_letter}{joint_итого_row})"
        return tcell

    I_TOTAL = _combine(I_TOTAL_TRACE, "F")
    J_TOTAL = _combine(J_TOTAL_TRACE, "G")
    K_TOTAL = _combine(K_TOTAL_TRACE, "H")
    L_TOTAL = _combine(L_TOTAL_TRACE, "I")
    M_TOTAL = _combine(M_TOTAL_TRACE, "J")

    label_font = Font(bold=True)

    def put(label: str, formula_or_value, units: str = "м³"):
        nonlocal r
        ws.cell(r, COL_BRANCH_ID).value = label
        ws.cell(r, COL_DN).value = units
        ws.cell(r, COL_TRACE_LEN).value = formula_or_value
        ws.cell(r, COL_BRANCH_ID).font = label_font
        r += 1
        return r - 1

    # Header for the summary section
    ws.cell(r, COL_BRANCH_ID).value = "СВОДКА ЗЕМЛЯНЫХ РАБОТ"
    ws.cell(r, COL_BRANCH_ID).font = Font(bold=True, size=12)
    ws.cell(r, COL_BRANCH_ID).fill = PatternFill("solid", fgColor="C6E0B4")
    r += 1

    # C53 snow = (I_total) * 3.2 * snow_thickness * IF(season_winter,1,0)
    snow_row = put(
        "Расчистка трассы от снега",
        f"=IF({_inp(IDX_SEASON_WINTER_FLAG)}=1,{I_TOTAL}*3.2*{_inp(IDX_SNOW_THICKNESS)},0)"
    )

    # C54 hand trench = hand_trench_count * length * width * depth
    # All four factors now live in ИсхДанные (B8, B9, B10, B1) — engineer edits
    # the panel, not the formula.
    hand_trench_row = put(
        "Разработка траншеи вручную (под пересечения)",
        f"={_inp(IDX_HAND_TRENCH_COUNT)}*{_inp(IDX_HAND_TRENCH_LEN)}"
        f"*{_inp(IDX_HAND_TRENCH_WIDTH)}*{_inp(IDX_DEPTH)}"
    )

    # C55 hand backfill = C54
    put("Обратная засыпка вручную", f"=C{hand_trench_row}")

    # C56 hand compaction = C55
    put("Уплотнение грунта обратной засыпки вручную", f"=C{hand_trench_row}")

    # C57 excavator = (K_total - C54) * 1.05
    excavator_row = put(
        "Разработка траншеи одноковшовым экскаватором (× 1.05)",
        f"=({K_TOTAL}-C{hand_trench_row})*1.05"
    )

    # C58 sand cushion total = M_total
    sand_cushion_row = put("Подушка из песка под трубу", f"={M_TOTAL}")

    # C59 sand sprinkle total = L_total + M_total*2
    sand_sprinkle_row = put("Присыпка из песка над трубой", f"={L_TOTAL}+{M_TOTAL}*2")

    # If exit_sand_total_cell_ref provided, include it in the excess-soil berm.
    exit_sand_term = f"+{exit_sand_total_cell_ref}" if exit_sand_total_cell_ref else ""

    # C62 excess = J_total + exit_sand + C58 + C59.
    # This soil is no longer hauled to TBO; it is formed into a berm above trench.
    landfill_row = put(
        "Устройство валика из вытесненного грунта",
        f"={J_TOTAL}{exit_sand_term}+C{sand_cushion_row}+C{sand_sprinkle_row}"
    )

    # C60 dozer backfill = C57 - C62
    dozer_row = put(
        "Обратная засыпка траншеи бульдозером",
        f"=C{excavator_row}-C{landfill_row}"
    )

    # C61 dozer compaction = C60
    put("Уплотнение бульдозерной засыпки", f"=C{dozer_row}")

    # Balance: D = sum(excavation), E = sum(backfill + berm excess), should be 0
    r += 1
    balance_label_row = r
    ws.cell(r, COL_BRANCH_ID).value = "БАЛАНС"
    ws.cell(r, COL_BRANCH_ID).font = Font(bold=True, size=12)
    ws.cell(r, COL_BRANCH_ID).fill = PatternFill("solid", fgColor="C6E0B4")
    r += 1
    ws.cell(r, COL_BRANCH_ID).value = "Выемка (ручная + экскаватор)"
    ws.cell(r, COL_TRACE_LEN).value = f"=C{hand_trench_row}+C{excavator_row}"
    excav_total_row = r
    r += 1
    ws.cell(r, COL_BRANCH_ID).value = "Засыпка (ручная + подушка + присыпка + бульдозер)"
    ws.cell(r, COL_TRACE_LEN).value = (
        f"=C{hand_trench_row}+C{sand_cushion_row}+C{sand_sprinkle_row}+C{dozer_row}"
    )
    backfill_total_row = r
    r += 1
    ws.cell(r, COL_BRANCH_ID).value = "Песок выходов (рециркуляция в засыпку)"
    # IMPORTANT: cell ref needs '=' prefix to be a formula, not literal text.
    # Without it Excel treats 'D17' as a string and the БАЛАНС row gets #ЗНАЧ!
    if exit_sand_total_cell_ref:
        ws.cell(r, COL_TRACE_LEN).value = f"={exit_sand_total_cell_ref}"
    else:
        ws.cell(r, COL_TRACE_LEN).value = 0
    exit_sand_recycled_row = r
    r += 1
    ws.cell(r, COL_BRANCH_ID).value = "Объём труб (занимает место)"
    ws.cell(r, COL_TRACE_LEN).value = f"={J_TOTAL}"
    pipe_recycled_row = r
    r += 1
    ws.cell(r, COL_BRANCH_ID).value = "Разница (должна = 0)"
    ws.cell(r, COL_BRANCH_ID).font = Font(bold=True)
    ws.cell(r, COL_TRACE_LEN).value = (
        f"=C{excav_total_row}-C{backfill_total_row}"
        f"-C{exit_sand_recycled_row}-C{pipe_recycled_row}"
    )
    ws.cell(r, COL_TRACE_LEN).font = Font(bold=True)
    balance_diff_row = r

    return {
        "snow":            f"C{snow_row}",
        "hand_trench":     f"C{hand_trench_row}",
        "excavator":       f"C{excavator_row}",
        "sand_cushion":    f"C{sand_cushion_row}",
        "sand_sprinkle":   f"C{sand_sprinkle_row}",
        "landfill":        f"C{landfill_row}",
        "dozer":           f"C{dozer_row}",
        "excav_total":     f"C{excav_total_row}",
        "backfill_total":  f"C{backfill_total_row}",
        "balance_diff":    f"C{balance_diff_row}",
    }


# ============================================================================
# Joint-pipe block (rendered between trace ИТОГО and exit-sand)
# ============================================================================

def build_joint_pipe_block(
    wb: Workbook, pairs: list[JointPipePair], trace_итого_row: int
) -> Optional[int]:
    """Write joint-pipe block on the trace sheet, just below trace ИТОГО.

    Layout (one row per JointPipePair plus header + ИТОГО):

      A=label B=DN1 C=DN2 D=max(B,C) E=length
      F=open_len(=E)  G=pipe_vol_J  H=soil_K  I=sand_side_L  J=sand_cushion_M

    Returns the row number of the joint ИТОГО (so summary can reference it),
    or None if pairs is empty.

    Trapezoidal trench with WIDER bottom (ИсхДанные!B${IDX_JOINT_BOTTOM_WIDTH},
    default 1.2 m) and depth based on max(DN1, DN2). Pipe volume sums both
    cross-sections: J = π × ((DN1/1000)² + (DN2/1000)²) × L / 4.
    """
    if not pairs:
        return None
    ws = wb[SHEET_TRACE]

    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    formula_fill = PatternFill("solid", fgColor="F2F2F2")
    итого_fill   = PatternFill("solid", fgColor="FFD966")

    r = trace_итого_row + 2   # blank row separator
    # Section title
    ws.cell(r, COL_BRANCH_ID).value = "СОВМЕСТНАЯ ПРОКЛАДКА"
    ws.cell(r, COL_BRANCH_ID).font = Font(bold=True, size=12)
    ws.cell(r, COL_BRANCH_ID).fill = PatternFill("solid", fgColor="C6E0B4")
    r += 1
    # Header row
    headers = [
        (COL_BRANCH_ID, "Id"),
        (COL_DN,         "Ду1, мм"),
        (COL_TRACE_LEN,  "Ду2, мм"),
        (COL_GNB_LEN,    "max(Ду1,Ду2)"),
        (COL_JOINT_LEN,  "Длина, м"),
        (COL_OPEN_LEN,   "I (= длина)"),
        (COL_PIPE_VOL,   "Объём труб J, м³"),
        (COL_SOIL_VOL,   "Объём грунта K, м³"),
        (COL_SAND_SIDE,  "Песок по сторонам L, м³"),
        (COL_SAND_CUSHION,"Подушка + присыпка M, м³"),
    ]
    for col, text in headers:
        c = ws.cell(r, col)
        c.value = text
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[r].height = 30
    r += 1

    first_row = r
    for pair in pairs:
        ws.cell(r, COL_BRANCH_ID).value = pair.id
        ws.cell(r, COL_DN).value = pair.dn1_mm
        ws.cell(r, COL_TRACE_LEN).value = pair.dn2_mm
        ws.cell(r, COL_GNB_LEN).value = f"=MAX(B{r},C{r})"
        ws.cell(r, COL_JOINT_LEN).value = round(pair.length_m, 3)
        # F (open) = E (length itself — joint is always open trench)
        ws.cell(r, COL_OPEN_LEN).value = f"=E{r}"
        # G = pipe_vol_J = PI() × ((DN1/1000)² + (DN2/1000)²) × L / 4
        ws.cell(r, COL_PIPE_VOL).value = (
            f"=PI()*((B{r}/1000)^2+(C{r}/1000)^2)*E{r}/4"
        )
        # H = soil_K — trapezoidal trench with JOINT bottom width
        # depth_total uses max DN (column D), bottom_width = IDX_JOINT_BOTTOM_WIDTH
        depth_total = (
            f"({_inp(IDX_DEPTH)}+D{r}/1000+IF({_inp(IDX_SAND_CUSHION_FLAG)}=1,{_inp(IDX_SAND_PAD)},0))"
        )
        area = f"({_inp(IDX_JOINT_BOTTOM_WIDTH)}+{depth_total}*{_inp(IDX_SLOPE_RATIO)})*{depth_total}"
        ws.cell(r, COL_SOIL_VOL).value = f"=F{r}*{area}"
        # I = sand_side L = IF(cushion, (max_DN/1000 * joint_bottom_width × E − G), 0)
        ws.cell(r, COL_SAND_SIDE).value = (
            f"=IF({_inp(IDX_SAND_CUSHION_FLAG)}=1,"
            f"(D{r}/1000*{_inp(IDX_JOINT_BOTTOM_WIDTH)}*E{r}-G{r}),0)"
        )
        # J = cushion M = IF(cushion, F × joint_bottom_width × sand_pad, 0)
        ws.cell(r, COL_SAND_CUSHION).value = (
            f"=IF({_inp(IDX_SAND_CUSHION_FLAG)}=1,"
            f"F{r}*{_inp(IDX_JOINT_BOTTOM_WIDTH)}*{_inp(IDX_SAND_PAD)},0)"
        )
        for col in (COL_GNB_LEN, COL_OPEN_LEN, COL_PIPE_VOL, COL_SOIL_VOL,
                    COL_SAND_SIDE, COL_SAND_CUSHION):
            ws.cell(r, col).fill = formula_fill
        r += 1
    last_row = r - 1

    # ИТОГО row
    ws.cell(r, COL_BRANCH_ID).value = "ИТОГО"
    ws.cell(r, COL_BRANCH_ID).font = Font(bold=True)
    for col in (COL_JOINT_LEN, COL_OPEN_LEN, COL_PIPE_VOL, COL_SOIL_VOL,
                COL_SAND_SIDE, COL_SAND_CUSHION):
        letter = get_column_letter(col)
        ws.cell(r, col).value = f"=SUM({letter}{first_row}:{letter}{last_row})"
        ws.cell(r, col).font = Font(bold=True)
        ws.cell(r, col).fill = итого_fill
    joint_итого = r
    return joint_итого


# ============================================================================
# Exit-sand block (rendered as inline section)
# ============================================================================

def append_exit_sand_block(
    wb: Workbook, nodes: list[ExitSandNode], итого_row: int
) -> str:
    """Append the «Песок для устройства выходов из земли» block below ИТОГО.

    Returns the cell reference to the total (e.g. "C{итого_row}+...").
    """
    ws = wb[SHEET_TRACE]

    # Find current bottom of sheet
    r = ws.max_row + 2
    ws.cell(r, COL_BRANCH_ID).value = "ПЕСОК НА УСТРОЙСТВО ВЫХОДОВ ИЗ ЗЕМЛИ"
    ws.cell(r, COL_BRANCH_ID).font = Font(bold=True, size=12)
    ws.cell(r, COL_BRANCH_ID).fill = PatternFill("solid", fgColor="C6E0B4")
    r += 1
    ws.cell(r, COL_BRANCH_ID).value = "Узел"
    ws.cell(r, COL_DN).value = "Кол-во"
    ws.cell(r, COL_TRACE_LEN).value = "Кратность × B7"
    ws.cell(r, COL_GNB_LEN).value = "Объём, м³"
    for c in range(COL_BRANCH_ID, COL_GNB_LEN + 1):
        ws.cell(r, c).font = Font(bold=True)
    r += 1

    first_node_row = r
    for node in nodes:
        ws.cell(r, COL_BRANCH_ID).value = node.name
        ws.cell(r, COL_DN).value = node.count
        ws.cell(r, COL_TRACE_LEN).value = node.sand_units
        ws.cell(r, COL_GNB_LEN).value = f"=B{r}*C{r}*{_inp(IDX_SAND_PER_GNB_EXIT)}"
        r += 1
    last_node_row = r - 1

    # ИТОГО
    ws.cell(r, COL_BRANCH_ID).value = "ИТОГО"
    ws.cell(r, COL_BRANCH_ID).font = Font(bold=True)
    ws.cell(r, COL_GNB_LEN).value = f"=SUM(D{first_node_row}:D{last_node_row})"
    ws.cell(r, COL_GNB_LEN).font = Font(bold=True)
    ws.cell(r, COL_GNB_LEN).fill = PatternFill("solid", fgColor="FFD966")
    total_cell = f"D{r}"
    return total_cell


# ============================================================================
# Public entry point
# ============================================================================

@dataclass
class XlsxRenderInputs:
    """Bundle for one xlsx render call."""
    project_params: ProjectParams
    branches: list[Branch]
    exit_sand_nodes: list[ExitSandNode]
    pits: Optional[PitParams] = None
    hand_trench_count: int = 0
    hand_trench_section_length: float = 4.0
    hand_trench_width: float = 0.7
    # Trench profile defaults (engineer can override in xlsx):
    bottom_width: float = 0.7
    slope_ratio: float = 0.0
    joint_bottom_width: float = 1.2
    pipe_overrun_factor: float = 1.10
    # Joint-pipe pairs (default empty — no joint laying in most projects).
    joint_pipe_pairs: list = field(default_factory=list)
    # Raw dwg blocks (for ВР node library auto-expansion). Empty = no ВР nodes section.
    blocks: list = field(default_factory=list)
    # Raw dwg polylines (for polyline-triggered templates, e.g. Узел А).
    polylines: list = field(default_factory=list)
    # Manual counts for templates with triggered_by="manual" (e.g. {"grpsh": 3}).
    manual_counts: dict = field(default_factory=dict)
    # Custom node library; None = use default shipped with pipe_engine.
    node_library: Optional[NodeLibrary] = None
    # Catalog bundle for VLOOKUP integration. If None, loaded from default pipe_catalog.db.
    catalog: Optional[CatalogBundle] = None
    # Source extract provenance for XML-shaped CVOR sheets.
    source_dwg_path: Optional[str] = None
    source_extracted_at_utc: Optional[str] = None
    sand_cushion_required: bool = False
    season_winter: bool = True


def write_xlsx(inputs: XlsxRenderInputs, output_path: str | Path) -> None:
    """Render the earthwork xlsx («Расчет земли» + ИсхДанные) to disk.

    All computed cells are formula strings. Caller (CLI or notebook)
    typically obtains a full ВОР document by post-processing this output
    alongside a catalog-lookup step — that integration lives elsewhere.
    """
    wb = Workbook()
    # Remove the default empty sheet — we'll create our own.
    default_sheet = wb.active
    wb.remove(default_sheet)

    build_inputs_sheet(
        wb, inputs.project_params,
        bottom_width_default=inputs.bottom_width,
        slope_ratio_default=inputs.slope_ratio,
        sand_cushion_required=inputs.sand_cushion_required,
        season_winter=inputs.season_winter,
        hand_trench_count=inputs.hand_trench_count,
        hand_trench_section_length=inputs.hand_trench_section_length,
        hand_trench_width=inputs.hand_trench_width,
        joint_bottom_width=inputs.joint_bottom_width,
        pipe_overrun_factor=inputs.pipe_overrun_factor,
    )

    first_data_row, итого_row = build_trace_sheet(
        wb, inputs.branches, joint_pair_count=len(inputs.joint_pipe_pairs)
    )

    # Joint-pipe block comes immediately after trace ИТОГО (so its addresses
    # match the forward-references that trace SUMIF formulas use).
    joint_итого_row = build_joint_pipe_block(wb, inputs.joint_pipe_pairs, итого_row)

    # Exit-sand block (returns total cell ref for use by summary)
    exit_sand_total_ref = None
    if inputs.exit_sand_nodes:
        exit_sand_total_ref = append_exit_sand_block(wb, inputs.exit_sand_nodes, итого_row)

    summary_layout = append_earthworks_summary(
        wb, итого_row,
        exit_sand_total_cell_ref=exit_sand_total_ref,
        pits=inputs.pits,
        joint_итого_row=joint_итого_row,
    )

    # Лист «ВР» (Ведомость объёмов работ) — главный итог.
    # Передаём layout map с semantic-references на нужные ячейки
    # Расчёта земли, чтобы ВР не зависел от конкретных row numbers.
    trace_layout = {
        "trace_итого_row": итого_row,
        "exit_sand_total_cell": exit_sand_total_ref or "0",
        "j_total_cell": f"G{итого_row}",
        "open_total_cell": f"F{итого_row}",
        "k_total_cell": f"H{итого_row}",
        **summary_layout,
    }
    # Catalog + Тег sheets — для VLOOKUP в ВР
    catalog = inputs.catalog
    if catalog is None:
        try:
            catalog = load_catalog()
        except FileNotFoundError:
            catalog = None  # no catalog → ВР emits literal tags
    build_vr_sheet(
        wb,
        branches=inputs.branches,
        blocks=inputs.blocks,
        node_library=inputs.node_library,
        trace_sheet_name=SHEET_TRACE,
        trace_layout=trace_layout,
        polylines=inputs.polylines,
        manual_counts=inputs.manual_counts,
        catalog=catalog,
    )
    build_spec_sheet_from_vr(wb, catalog=catalog)
    build_cvor_sheets(
        wb,
        catalog=catalog,
        source_dwg_path=inputs.source_dwg_path,
        source_extracted_at_utc=inputs.source_extracted_at_utc,
    )

    # Write
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))


# ============================================================================
# Каталог + Тег sheets
# ============================================================================

def build_catalog_sheet(wb: Workbook, catalog: CatalogBundle) -> None:
    """Write Каталог sheet — 995 rows × 13 columns, mirrors reference structure.

    Column order is fixed (used by VLOOKUP col-index references in ВР):
      A=DN  B=DN2  C=Поз  D=Наименование  E=Тип/марка  F=Код  G=Завод
      H=Ед.изм  I=Кол-во  J=Масса  K=Примеч.  L=Тэги  M=Поз(=ROW)
    """
    ws = wb.create_sheet(SHEET_CATALOG)
    headers = ["DN", "DN2", "Поз", "Наименование и техническая характеристика",
               "Тип, марка, обозначение", "Код оборудования", "Завод-изготовитель",
               "Ед.изм.", "Кол-во", "Масса, кг", "Примеч.", "Тэги", "Поз"]
    hf = Font(bold=True)
    hfill = PatternFill("solid", fgColor="D9E1F2")
    for i, h in enumerate(headers, start=1):
        c = ws.cell(1, i)
        c.value = h
        c.font = hf
        c.fill = hfill
    # Rows: id from SQLite = original Excel row number; we offset by 1 since
    # header is on row 1. Data row r = sqlite_id (so VLOOKUP ranges match
    # what tag_pivot.min_pos/max_pos return).
    for cr in catalog.rows:
        r = cr.row_id
        ws.cell(r, 1).value = _try_int(cr.dn)
        ws.cell(r, 2).value = cr.dn2
        ws.cell(r, 3).value = cr.pos
        ws.cell(r, 4).value = cr.name
        ws.cell(r, 5).value = cr.type_mark
        ws.cell(r, 6).value = cr.code
        ws.cell(r, 7).value = cr.manufacturer
        ws.cell(r, 8).value = cr.unit
        ws.cell(r, 9).value = cr.qty
        ws.cell(r, 10).value = cr.mass_per_unit
        ws.cell(r, 11).value = cr.note
        ws.cell(r, 12).value = cr.tag
        ws.cell(r, 13).value = r   # literal pos, не формула (для test eval)
    # Column widths
    widths = {"A":6,"B":8,"C":6,"D":50,"E":30,"F":20,"G":25,"H":8,"I":8,
              "J":8,"K":15,"L":25,"M":6}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def build_tag_sheet(wb: Workbook, catalog: CatalogBundle) -> None:
    """Write Тег sheet — pivot tag → (min_pos, max_pos)."""
    ws = wb.create_sheet(SHEET_TAG)
    ws.cell(1, 1).value = "Названия строк"
    ws.cell(1, 2).value = "Минимум по полю Поз"
    ws.cell(1, 3).value = "Максимум по полю Поз"
    for c in (1, 2, 3):
        ws.cell(1, c).font = Font(bold=True)
        ws.cell(1, c).fill = PatternFill("solid", fgColor="D9E1F2")
    for i, t in enumerate(catalog.tags, start=2):
        ws.cell(i, 1).value = t.tag
        ws.cell(i, 2).value = t.min_pos
        ws.cell(i, 3).value = t.max_pos
    ws.column_dimensions["A"].width = 30


def _try_int(s):
    """Try to convert DN string to int (e.g. '110' → 110). Leave as TEXT for '63x57'."""
    if s is None:
        return None
    try:
        return int(s)
    except (ValueError, TypeError):
        return s
