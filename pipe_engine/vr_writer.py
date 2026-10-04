"""
vr_writer.py — Лист «ВР» (Ведомость объёмов работ) для xlsx-выхода.

ВР это **главный итоговый лист** проектной документации, который сметчик
получает на руки. Структура (как в эталонной ВР):

  Section 1: Земляные работы (15 пунктов из «Расчёт земли»)
  Section 2: Общая длина труб (per DN, ROUNDUP с overrun-factor)
  Section 3: Укладка открытым способом (per DN)
  Section 4: Укладка ГНБ (per DN с gnb_length > 0)
  Section 5: Узлы (auto-emit из NodeLibrary для блоков СО_* которые
              match template — КТ, ЗУ_кран, и т.д.)

Каждая section_row — section_header + N item rows. Иерархическая
нумерация через Excel-формулы (как в эталоне row 5+):
  A: =IF(MAX($A$3:A4)=0,1,MAX($A$3:A4)+1)
  B: subnumber
  C: =MAX(A$3:An) & "." & Bn

Section 5 группирует blocks по (template_id, host_dn) — иначе для
В эталоне получили бы 800+ строк (385 КТ + 321 Узел_А + ...). Группировка
даёт ~10-20 «узел-групп» × N компонентов.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .catalog_loader import CatalogBundle, CatalogRow, WorkCatalogRow
from .io_schema import BlockRecord, PolylineRecord
from .node_library import NodeLibrary, NodeTemplate
from .trench_profile import Branch


SHEET_CATALOG_NAME = "Каталог"
CATALOG_COL_NAME_IDX = 4   # «Наименование» — column index в Каталог sheet


SHEET_VR = "ВР"

# Column layout
COL_SECT_NUM   = 1   # A — section number (1, 2, 3, ...)
COL_ITEM_NUM   = 2   # B — item number within section (1, 2, ...)
COL_HIER       = 3   # C — hierarchical "1.1", "1.2" (Excel formula)
COL_NAME       = 4   # D — Наименование работ
COL_UNIT       = 5   # E — Ед. изм.
COL_QTY        = 6   # F — Кол-во (formula)
COL_REF        = 7   # G — Ссылка на чертежи/спецификацию
COL_FORMULA    = 8   # H — Формула расчёта (текст для аудита)
COL_NOTE       = 9   # I — Примечание

HEADER_ROW = 3      # column titles
FIRST_DATA_ROW = 5  # first section starts here (matches reference row 5)

OVERRUN_FACTOR_FORMULA_REF = "ИсхДанные!$B$13"   # see xlsx_writer IDX_PIPE_OVERRUN_FACTOR


# ============================================================================
# Public entry point
# ============================================================================

def build_vr_sheet(
    wb: Workbook,
    branches: list[Branch],
    blocks: list[BlockRecord],
    node_library: Optional[NodeLibrary] = None,
    *,
    trace_sheet_name: str = "Расчет земли",
    trace_layout: Optional[dict] = None,
    polylines: Optional[list[PolylineRecord]] = None,
    manual_counts: Optional[dict[str, int]] = None,
    catalog: Optional[CatalogBundle] = None,
) -> int:
    """Render the «ВР» sheet. Returns row count for diagnostics.

    ``trace_layout`` is a dict of semantic cell references on the trace sheet:
      snow, hand_trench, excavator, sand_cushion, sand_sprinkle, landfill,
      dozer, excav_total, backfill_total, balance_diff, exit_sand_total_cell,
      j_total_cell, open_total_cell, k_total_cell.
    Built by xlsx_writer.write_xlsx and passed in. If None, defaults to a
    layout matching the 5-row reference-snapshot test fixture (for unit tests).
    """
    if node_library is None:
        node_library = NodeLibrary.load_default()
    if trace_layout is None:
        # Fallback for tests / direct callers: layout for 5-pipe reference snapshot
        trace_layout = _default_trace_layout_5_pipes()

    ws = wb.create_sheet(SHEET_VR)
    _write_header(ws)
    r = FIRST_DATA_ROW

    # Section 1: Земляные работы
    r = _write_section_zemlanye(ws, r, trace_sheet_name, trace_layout, catalog=catalog)

    # Section 2: Общая длина труб
    r = _write_section_total_pipes(ws, r, branches, trace_sheet_name, catalog=catalog)

    # Section 3: Укладка открытым способом
    r = _write_section_open_laying(ws, r, branches, trace_sheet_name, catalog=catalog)

    # Section 4: Укладка ГНБ (только DN где gnb_length > 0)
    r = _write_section_gnb_laying(ws, r, branches, trace_sheet_name, catalog=catalog)

    # Section 5: Узлы (через NodeLibrary — block + polyline + manual triggers)
    r = _write_section_nodes(ws, r, blocks, node_library,
                             polylines=polylines or [],
                             manual_counts=manual_counts or {},
                             catalog=catalog)

    _set_column_widths(ws)
    return r - 1


def _default_trace_layout_5_pipes() -> dict:
    """Hardcoded layout for the 5-pipe reference-snapshot fixture (legacy tests)."""
    return {
        "snow":             "C19",
        "hand_trench":      "C20",
        "excavator":        "C23",
        "sand_cushion":     "C24",
        "sand_sprinkle":    "C25",
        "landfill":         "C26",
        "dozer":            "C27",
        "excav_total":      "C31",
        "backfill_total":   "C32",
        "balance_diff":     "C35",
        "exit_sand_total_cell": "D16",
        "j_total_cell":     "G7",
        "open_total_cell":  "F7",
        "k_total_cell":     "H7",
    }


# ============================================================================
# Header / styling
# ============================================================================

def _write_header(ws) -> None:
    ws.cell(1, COL_NAME).value = "Ведомость объёмов работ"
    ws.cell(1, COL_NAME).font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, end_row=1, start_column=COL_NAME, end_column=COL_NOTE)
    headers = [
        (COL_SECT_NUM, "№"), (COL_ITEM_NUM, "п/п"),
        (COL_HIER, "№"), (COL_NAME, "Наименование работ"),
        (COL_UNIT, "Ед. изм."), (COL_QTY, "Кол-во"),
        # Hidden audit columns (unhide для инженерной проверки):
        (COL_REF, "Ссылка (debug)"),
        (COL_FORMULA, "Формула (debug)"),
        (COL_NOTE, "Метаданные (debug)"),
    ]
    hf = Font(bold=True)
    hfill = PatternFill("solid", fgColor="D9E1F2")
    for col, text in headers:
        c = ws.cell(HEADER_ROW, col)
        c.value = text
        c.font = hf
        c.fill = hfill
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[HEADER_ROW].height = 28


def _set_column_widths(ws) -> None:
    # Видимые сметчику колонки: A..F (№, п/п, иерархия, имя, ед., кол-во).
    # G/H/I — debug/audit (ссылка на исх. ячейку, формула расчёта, метаданные
    # NodeLibrary/VLOOKUP). Скрываем по умолчанию: инженер делает Unhide
    # когда нужно проверить происхождение числа.
    widths = {"A": 5, "B": 5, "C": 7, "D": 60, "E": 8, "F": 12,
              "G": 14, "H": 26, "I": 18}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for hidden_col in ("G", "H", "I"):
        ws.column_dimensions[hidden_col].hidden = True


def _write_section_header(ws, r: int, label: str) -> int:
    """Write section header row (only A and D filled). Returns next row."""
    ws.cell(r, COL_SECT_NUM).value = (
        f"=IF(MAX($A$3:A{r-1})=0,1,MAX($A$3:A{r-1})+1)"
    )
    ws.cell(r, COL_NAME).value = label
    ws.cell(r, COL_NAME).font = Font(bold=True)
    ws.cell(r, COL_NAME).fill = PatternFill("solid", fgColor="FFE699")
    return r + 1


def _write_item_row(
    ws, r: int, subnum: int,
    name: str | None, unit: str, qty: str | float | int,
    ref: str = "", formula_text: str = "", note: str = "",
) -> int:
    """Write one work-item row. Returns next row."""
    ws.cell(r, COL_ITEM_NUM).value = subnum
    ws.cell(r, COL_HIER).value = (
        f'=MAX($A$3:A{r}) & "." & B{r}'
    )
    if name is not None:
        ws.cell(r, COL_NAME).value = name
    ws.cell(r, COL_UNIT).value = unit
    ws.cell(r, COL_QTY).value = qty
    if ref:
        ws.cell(r, COL_REF).value = ref
    if formula_text:
        ws.cell(r, COL_FORMULA).value = formula_text
    if note:
        ws.cell(r, COL_NOTE).value = note
    return r + 1


# ============================================================================
# Section 1: Земляные работы
# ============================================================================

def _work_row_from_db(catalog: Optional[CatalogBundle], work_key: str) -> Optional[WorkCatalogRow]:
    if catalog is None:
        return None
    return catalog.work_by_key(work_key)


def _work_note(work_key: str, row: Optional[WorkCatalogRow]) -> str:
    if row is None:
        return f"work_key={work_key}, work_catalog=missing"
    norm = row.estimate_norm_code or "pending"
    return f"work_key={work_key}, estimate_norm_code={norm}, source={row.source or ''}"


def _write_section_zemlanye(
    ws,
    r: int,
    trace_sheet: str,
    layout: dict,
    *,
    catalog: Optional[CatalogBundle] = None,
) -> int:
    r = _write_section_header(ws, r, "Земляные работы:")
    # ВАЖНО: «Объём уложенных труб (для баланса)» намеренно НЕ включаем —
    # это служебная строка для проверки баланса в «Расчёте земли», не работа.
    # Audit-комментарий = semantic-ключ layout-map (легко сверить с trace).
    items = [
        # (work_key, fallback_name, fallback_unit, layout_key, formula_text)
        ("earth.snow_clearance", "Расчистка трассы от снега", "м3", "snow", "Расчет земли → snow"),
        ("earth.excavator_trench", "Разработка грунта в траншеях одноковшовым экскаватором", "м3", "excavator", "Расчет земли → excavator (K_итого − ручная) × 1.05"),
        ("earth.manual_crossing_trench", "Разработка грунта вручную в местах пересечения с инженерными коммуникациями", "м3", "hand_trench", "Расчет земли → hand_trench"),
        ("earth.manual_backfill", "Обратная засыпка вручную", "м3", "hand_trench", "Расчет земли → hand_trench (= объём ручной разработки)"),
        ("earth.manual_compaction", "Уплотнение грунта обратной засыпки вручную", "м3", "hand_trench", "Расчет земли → hand_trench"),
        ("earth.dozer_backfill", "Обратная засыпка траншеи бульдозером", "м3", "dozer", "Расчет земли → dozer"),
        ("earth.mechanical_compaction", "Механическое уплотнение грунта обратной засыпки", "м3", "dozer", "Расчет земли → dozer"),
        ("earth.excess_soil_berm", "Устройство валика из вытесненного грунта", "м3", "landfill", "Расчет земли → excess soil berm"),
        ("earth.exit_sand", "Песок для устройства выходов из земли", "м3", "exit_sand_total_cell", "Расчет земли → exit_sand_total"),
        ("earth.sand_bedding", "Устройство песчаного основания", "м3", "sand_cushion", "Расчет земли → sand_cushion"),
        ("earth.sand_sprinkle", "Песчаная присыпка над газопроводом", "м3", "sand_sprinkle", "Расчет земли → sand_sprinkle"),
    ]
    for i, (work_key, fallback_name, fallback_unit, layout_key, formula_text) in enumerate(items, start=1):
        work = _work_row_from_db(catalog, work_key)
        name = work.name_for_vor if work is not None else fallback_name
        unit = work.unit if work is not None else fallback_unit
        cell = layout[layout_key]
        # cell может быть "0" (когда exit_sand отсутствует) — тогда literal
        if cell == "0":
            qty_formula = 0
            ref = ""
        else:
            qty_formula = f"='{trace_sheet}'!{cell}"
            ref = f"{trace_sheet}!{cell}"
        r = _write_item_row(ws, r, i, name, unit, qty_formula,
                            ref=ref, formula_text=formula_text,
                            note=_work_note(work_key, work))
    return r + 1   # blank row separator


# ============================================================================
# Section 2: Общая длина труб (per DN with overrun factor)
# ============================================================================

def _branch_labels(branches: list[Branch]) -> dict[str, str]:
    """Map branch.id → user-facing label for disambiguation in pipe rows.

    При одной ветке — пустые labels (suffix не нужен).
    При нескольких — сначала пробуем ``section.name``. Если у разных
    branches одинаковый section.name (например все «cli_sand_winter»
    из CLI defaults) — для коллидирующих fallback на branch.id, иначе
    сметчик увидит «Труба DN63» × 2 без разницы.
    """
    if len(branches) <= 1:
        return {b.id: "" for b in branches}
    raw = {}
    for b in branches:
        raw[b.id] = b.section.name if b.section and b.section.name else b.id
    # Подсчитываем коллизии — если name использован более чем одной веткой,
    # тот name неинформативен и каждой коллидирующей выдаём её branch.id.
    counts: dict[str, int] = {}
    for n in raw.values():
        counts[n] = counts.get(n, 0) + 1
    labels: dict[str, str] = {}
    for bid, name in raw.items():
        labels[bid] = bid if counts[name] > 1 else name
    return labels


def _pipe_row_name(dn_mm: int, branch: Branch, branch_label: str,
                   *, prefix: str = "Труба ", suffix: str = "") -> str:
    """Construct pipe row name. ``branch_label`` пустой → без суффикса."""
    base = f"{prefix}DN{dn_mm}{suffix}"
    if not branch_label:
        return base
    return f"{base}, {branch_label}"


def _trace_row_for(branch_idx: int, pipe_idx: int, branches: list[Branch]) -> int:
    """Compute the trace-sheet row number for a given (branch, pipe).

    Mirrors xlsx_writer build_trace_sheet layout: TRACE_FIRST_DATA_ROW=2,
    rows are flat-ordered by branch then by pipe within branch.
    """
    TRACE_FIRST = 2
    offset = 0
    for bi, b in enumerate(branches):
        if bi < branch_idx:
            offset += len(b.pipes)
        elif bi == branch_idx:
            return TRACE_FIRST + offset + pipe_idx
    raise IndexError(f"branch_idx {branch_idx} out of range")


def _write_section_total_pipes(
    ws,
    r: int,
    branches: list[Branch],
    trace_sheet: str,
    *,
    catalog: Optional[CatalogBundle] = None,
) -> int:
    r = _write_section_header(ws, r, "Общая протяжённость укладываемых труб (с запасом ×B13):")
    subnum = 1
    sum_formula_parts: list[str] = []
    labels = _branch_labels(branches)
    for bi, b in enumerate(branches):
        for pi, p in enumerate(b.pipes):
            trace_row = _trace_row_for(bi, pi, branches)
            # F = open length on trace sheet (column F per xlsx_writer COL_OPEN_LEN=6)
            qty = f"=ROUNDUP('{trace_sheet}'!F{trace_row}*{OVERRUN_FACTOR_FORMULA_REF},0)"
            catalog_row = _catalog_row_from_db(catalog, "Труба ПЭ", p.dn_mm) if catalog is not None else None
            name = catalog_row.name if catalog_row is not None else _pipe_row_name(p.dn_mm, b, labels[b.id])
            note = (
                f"DB lookup: tag='Труба ПЭ', DN={p.dn_mm}, catalog_id={catalog_row.row_id}"
                if catalog_row is not None else ""
            )
            r = _write_item_row(ws, r, subnum, name, "м", qty,
                                ref=f"{trace_sheet}!F{trace_row}",
                                formula_text=f"ROUNDUP(open × overrun, 0)",
                                note=note)
            sum_formula_parts.append(f"F{r-1}")
            subnum += 1
    # Add summary row at top of section: total length sum
    # (We've already advanced — insert summary AFTER items as a "ИТОГО" row)
    if sum_formula_parts:
        r = _write_item_row(ws, r, subnum, "ИТОГО общая длина труб", "м",
                            f"=SUM({sum_formula_parts[0]}:{sum_formula_parts[-1]})",
                            formula_text="SUM of pipe rows above")
        ws.cell(r-1, COL_NAME).font = Font(bold=True)
        ws.cell(r-1, COL_QTY).font = Font(bold=True)
    return r + 1


# ============================================================================
# Section 3: Укладка открытым способом
# ============================================================================

def _work_pipe_name(
    catalog: Optional[CatalogBundle],
    work_key: str,
    dn_mm: int,
    branch_label: str,
    fallback: str,
) -> tuple[str, str, str]:
    work = _work_row_from_db(catalog, work_key)
    base = work.name_for_vor if work is not None else fallback
    unit = work.unit if work is not None else "м"
    name = f"{base} DN{dn_mm}"
    if branch_label:
        name = f"{name}, {branch_label}"
    return name, unit, _work_note(work_key, work)


def _write_section_open_laying(
    ws,
    r: int,
    branches: list[Branch],
    trace_sheet: str,
    *,
    catalog: Optional[CatalogBundle] = None,
) -> int:
    r = _write_section_header(ws, r, "Укладка газопроводов подземно открытым способом:")
    subnum = 1
    labels = _branch_labels(branches)
    for bi, b in enumerate(branches):
        for pi, p in enumerate(b.pipes):
            # Open laying = trace - GNB (we already have it as F in trace sheet, which is C-D-E)
            trace_row = _trace_row_for(bi, pi, branches)
            qty = f"='{trace_sheet}'!F{trace_row}"
            name, unit, note = _work_pipe_name(
                catalog,
                "pipe.pe_open_laying",
                p.dn_mm,
                labels[b.id],
                "Укладка газопровода из полиэтиленовых труб открытым способом",
            )
            r = _write_item_row(ws, r, subnum, name, "м", qty,
                                ref=f"{trace_sheet}!F{trace_row}",
                                formula_text="open trench length",
                                note=note)
            subnum += 1
    return r + 1


# ============================================================================
# Section 4: Укладка ГНБ
# ============================================================================

def _write_section_gnb_laying(
    ws,
    r: int,
    branches: list[Branch],
    trace_sheet: str,
    *,
    catalog: Optional[CatalogBundle] = None,
) -> int:
    has_any_gnb = any(p.gnb_length > 0 for b in branches for p in b.pipes)
    if not has_any_gnb:
        return r   # skip section entirely
    r = _write_section_header(ws, r, "Укладка газопроводов подземно закрытым способом (ГНБ):")
    subnum = 1
    labels = _branch_labels(branches)
    for bi, b in enumerate(branches):
        for pi, p in enumerate(b.pipes):
            if p.gnb_length <= 0:
                continue
            trace_row = _trace_row_for(bi, pi, branches)
            # D column = gnb_length on trace sheet
            qty = f"='{trace_sheet}'!D{trace_row}"
            name, unit, note = _work_pipe_name(
                catalog,
                "pipe.pe_gnb_laying",
                p.dn_mm,
                labels[b.id],
                "Прокладка газопровода из полиэтиленовых труб закрытым способом методом ГНБ",
            )
            r = _write_item_row(ws, r, subnum, name, unit, qty,
                                ref=f"{trace_sheet}!D{trace_row}",
                                formula_text="GNB length from extract",
                                note=note)
            subnum += 1
    return r + 1


# ============================================================================
# Section 5: Узлы (via NodeLibrary)
# ============================================================================

@dataclass
class _NodeGroup:
    template: NodeTemplate
    host_dn: Optional[int]
    block_count: int
    blocks: list[BlockRecord] = field(default_factory=list)


def _group_polylines_by_template(
    polylines: list[PolylineRecord], lib: NodeLibrary,
) -> list[_NodeGroup]:
    """For polyline-triggered templates: count matching polylines per (template, host_dn).

    host_dn is extracted from the layer name via regex capture group named 'dn'
    (e.g. for layer "!!ГСН3_Г1_ГП_Вводы@63" with pattern containing "@(?P<dn>\\d+)$"
    → host_dn=63).
    """
    import re
    buckets: dict[tuple[str, Optional[int]], int] = defaultdict(int)
    for tmpl in lib.templates_by_trigger("polyline_layer"):
        if not tmpl.polyline_layer_pattern:
            continue
        regex = re.compile(tmpl.polyline_layer_pattern)
        for p in polylines:
            m = regex.match(p.layer)
            if not m:
                continue
            try:
                host_dn = int(m.group("dn"))
            except (IndexError, ValueError):
                host_dn = None
            buckets[(tmpl.id, host_dn)] += 1
    out: list[_NodeGroup] = []
    for (tmpl_id, dn), count in sorted(buckets.items(), key=lambda x: (x[0][0], x[0][1] or 0)):
        out.append(_NodeGroup(template=lib.find_by_id(tmpl_id), host_dn=dn, block_count=count))
    return out


def _group_manual_templates(
    manual_counts: dict[str, int], lib: NodeLibrary,
) -> list[_NodeGroup]:
    """For manual-triggered templates: take count from CLI-supplied dict."""
    out: list[_NodeGroup] = []
    for tmpl in lib.templates_by_trigger("manual"):
        key = tmpl.manual_count_key
        if not key or key not in manual_counts:
            continue
        count = manual_counts[key]
        if count <= 0:
            continue
        out.append(_NodeGroup(template=tmpl, host_dn=None, block_count=count))
    return out


def _group_blocks_by_template(
    blocks: list[BlockRecord], lib: NodeLibrary,
    *,
    polylines: Optional[list[PolylineRecord]] = None,
) -> list[_NodeGroup]:
    """Group blocks by (template, host_dn).

    host_dn resolution per template.host_dn_source:
      * ``"attribute:<TAG>"`` — parse DN из named ATTRIB на блоке
        (e.g. "Седло DN110х63" → 110). Самый надёжный.
      * ``"geometry"`` — для блоков без DN-атрибута (КТ, ЗУ-кран):
        находим ближайшую trace-полилинию к ``block.position`` и берём
        DN из имени слоя (e.g. ``!!ГСН3_Г1_ГП@110`` → 110). Требует
        ``polylines`` чтобы было что искать. Closes Q5.
      * ``"layer"`` / ``"manual"`` / другое / unparseable → host_dn=None.
    """
    from .interpret import parse_dn_from_attr, resolve_geometric_host_dn

    buckets: dict[tuple[str, Optional[int]], list[BlockRecord]] = defaultdict(list)
    for b in blocks:
        tmpl = lib.find_by_block(b.block_name_effective)
        if tmpl is None:
            continue
        host_dn: Optional[int] = None
        src = tmpl.host_dn_source
        if src.startswith("attribute:"):
            attr_name = src.split(":", 1)[1]
            text = b.attributes.get(attr_name, "")
            dn_pair = parse_dn_from_attr(text)
            host_dn = dn_pair[0]
            # Fallback на geometry если атрибут не дал DN (ЗУ-кран часто
            # подписан "Кран КШ-150с" — без явного DN — берём с трассы).
            if host_dn is None and polylines:
                host_dn = resolve_geometric_host_dn(b.position, polylines)
        elif src == "geometry" and polylines:
            host_dn = resolve_geometric_host_dn(b.position, polylines)
        # 'layer' / 'manual' / unparseable → host_dn stays None
        buckets[(tmpl.id, host_dn)].append(b)

    out: list[_NodeGroup] = []
    for (tmpl_id, dn), group_blocks in sorted(buckets.items(), key=lambda x: (x[0][0], x[0][1] or 0)):
        out.append(_NodeGroup(template=lib.find_by_id(tmpl_id),
                              host_dn=dn, block_count=len(group_blocks),
                              blocks=group_blocks))
    return out


def _parse_float_attr(text: str) -> Optional[float]:
    if not text:
        return None
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return None


def _fmt_num(x: float) -> str:
    return f"{x:.6f}".rstrip("0").rstrip(".")


def _onk_step_for_host_dn(host_dn: Optional[int]) -> float:
    # Draft rule from futlyar_pe.json notes: B=2.5m for OD <=300mm,
    # 3.5m for 400-600mm, 5.5m above that.
    if host_dn is None or host_dn <= 300:
        return 2.5
    if host_dn <= 600:
        return 3.5
    return 5.5


def _futlyar_lengths(blocks: list[BlockRecord]) -> list[float]:
    values: list[float] = []
    for b in blocks:
        length = _parse_float_attr(b.attributes.get("LENGTH", ""))
        if length is not None and length > 0:
            values.append(length)
    return values


def _node_component_qty_override(
    tmpl: NodeTemplate,
    comp_tag: str,
    grp: _NodeGroup,
) -> Optional[tuple[str, str]]:
    """Return (qty_formula, audit_text) for per-instance node quantities."""
    if tmpl.id != "futlyar_pe":
        return None

    lengths = _futlyar_lengths(grp.blocks)
    if not lengths:
        return None

    total_length = sum(lengths)
    length_formula = f"={_fmt_num(total_length)}"
    length_audit = " + ".join(_fmt_num(v) for v in lengths)

    tag_lower = comp_tag.lower()
    if "труба-футляр" in tag_lower or "прокладка футляра" in tag_lower:
        return (length_formula, f"LENGTH attrs: {length_audit}")

    if "кольцо опорно" in tag_lower or "установка опорно" in tag_lower:
        step = _onk_step_for_host_dn(grp.host_dn)
        terms = [f"ROUND({_fmt_num(v)}/{_fmt_num(step)},0)+1" for v in lengths]
        return (f"={'+'.join(terms)}", f"ONK per футляр: {' + '.join(terms)}")

    return None


def _catalog_row_from_db(catalog: CatalogBundle, tag_substring: str,
                         host_dn: Optional[int]) -> Optional[CatalogRow]:
    """Resolve catalog row from SQLite; no workbook Каталог sheet required."""
    if host_dn is None:
        return None
    direct = catalog.find_row_by_tag_and_dn(tag_substring, host_dn)
    if direct is not None:
        return direct
    needle = tag_substring.lower()
    for row in catalog.rows:
        if row.tag and needle in row.tag.lower():
            return row
    return None


def _write_section_nodes(ws, r: int, blocks: list[BlockRecord],
                          lib: NodeLibrary,
                          polylines: Optional[list[PolylineRecord]] = None,
                          manual_counts: Optional[dict[str, int]] = None,
                          catalog: Optional[CatalogBundle] = None) -> int:
    block_groups = _group_blocks_by_template(blocks, lib, polylines=polylines or [])
    poly_groups = _group_polylines_by_template(polylines or [], lib)
    manual_groups = _group_manual_templates(manual_counts or {}, lib)
    groups = block_groups + poly_groups + manual_groups
    if not groups:
        return r
    r = _write_section_header(ws, r, "Устройство узлов:")
    subnum = 1
    for grp in groups:
        tmpl = grp.template
        variant = tmpl.pick_variant(grp.host_dn) if grp.host_dn else None
        # Видимое сметчику имя: «<имя узла> DN110, в т. ч.:» если DN известен,
        # иначе просто «<имя узла>, в т. ч.:» без шумного «DN не определён».
        dn_str = f" DN{grp.host_dn}" if grp.host_dn else ""
        # «(УГ DN50)» suffix скрыт — это внутренний variant код типового
        # чертежа (УГ 26.00 СБ), сметчику он не нужен и сбивает с толку
        # рядом с уже показанным DN газопровода. Сохраняем variant info в
        # hidden audit-col I (для инженерной проверки).
        suffix_str = ""
        # ----- HEADER ROW (категорийный, НЕ-sumable) ----------------------
        # Чтобы убрать визуальный двойной счёт «узел: 385 / компонент: 385»,
        # header строки получает: НИКАКОЙ B/C нумерации, пустые E/F (сметчик
        # не суммирует header в позиции), block_count — только в имени.
        # Стиль grey-italic делает header визуально отличным от sumable rows.
        block_count = grp.block_count
        ws.cell(r, COL_NAME).value = (
            f"{tmpl.name}{dn_str}{suffix_str} "
            f"(всего {block_count} шт.), в т. ч.:"
        )
        ws.cell(r, COL_NAME).font = Font(italic=True, bold=True)
        ws.cell(r, COL_NAME).fill = PatternFill("solid", fgColor="EAEAEA")
        # B, C, E, F остаются пустыми → header не попадает в авто-сумму сметчика
        ws.cell(r, COL_REF).value = tmpl.typical_drawing_ref or ""
        variant_note = (f", variant={variant.suffix}"
                        if variant and variant.suffix else "")
        ws.cell(r, COL_NOTE).value = (
            f"template={tmpl.id}, host_dn={grp.host_dn}, "
            f"blocks={block_count}{variant_note}"
        )
        r += 1
        # subnum НЕ инкрементируем — header не позиция
        # ----- COMPONENT ROWS (sumable, нумеруются) -----------------------
        for comp_idx, comp in enumerate(tmpl.components, start=1):
            # qty: Excel-формула с literal block_count в формуле.
            # block_count приходит из dwg (immutable для xlsx), поэтому literal —
            # не теряем audit-ability, но избегаем зависимости от header F.
            # Когда block_count=1 ИЛИ qty=1 — упрощаем до literal числа
            # (нет смысла в формуле «=1*N» или «=N*1»).
            qty_override = _node_component_qty_override(tmpl, comp.tag, grp)
            if qty_override is not None:
                comp_qty_formula, audit_text = qty_override
            elif isinstance(comp.qty, str):
                comp_qty_formula = comp.qty               # custom Excel-формула из шаблона
                audit_text = str(comp.qty)
            elif float(comp.qty).is_integer():
                k = int(comp.qty)
                if k == 1:
                    comp_qty_formula = block_count        # literal
                elif block_count == 1:
                    comp_qty_formula = k                  # literal
                else:
                    comp_qty_formula = f"={block_count}*{k}"
                audit_text = f"{k} × {block_count} (block_count)"
            else:
                if block_count == 1:
                    # ROUND(1×0.6, 1) = 0.6 — но округление всё равно нужно
                    # т.к. qty=0.287 → 0.3, не литерал 0.287. Используем ROUND.
                    comp_qty_formula = f"=ROUND({comp.qty},1)"
                else:
                    comp_qty_formula = f"=ROUND({block_count}*{comp.qty},1)"
                audit_text = f"ROUND({comp.qty} × {block_count}, 1)"

            ws.cell(r, COL_ITEM_NUM).value = subnum
            ws.cell(r, COL_HIER).value = f'=MAX($A$3:A{r}) & "." & B{r}'

            # Catalog DB lookup — если есть match и catalog загружен.
            # dn_hint компонента перебивает host_dn узла (multi-DN узлы:
            # Узел А, Узел Б, ГРПШ, Врезка — где футляр/трубка/отвод
            # имеют свой DN отличный от основного газопровода).
            cm_tag = (comp.catalog_match or {}).get("tag")
            effective_dn = comp.dn_hint if comp.dn_hint is not None else grp.host_dn
            catalog_row = None
            if catalog is not None and cm_tag:
                catalog_row = _catalog_row_from_db(catalog, cm_tag, effective_dn)
            if catalog_row is not None:
                ws.cell(r, COL_NAME).value = catalog_row.name or f"  • {comp.tag}"
                ws.cell(r, COL_NOTE).value = (
                    f"DB lookup: tag='{cm_tag}', DN={effective_dn}, "
                    f"catalog_id={catalog_row.row_id}"
                    + (f" (dn_hint override of host_dn={grp.host_dn})"
                       if comp.dn_hint is not None else "")
                )
            else:
                # Fallback: literal tag (нет catalog или нет matching tag-range)
                ws.cell(r, COL_NAME).value = f"  • {comp.tag}"
                if comp.catalog_match:
                    cm = comp.catalog_match
                    ws.cell(r, COL_NOTE).value = f"catalog: {cm} (no VLOOKUP range)"

            ws.cell(r, COL_UNIT).value = comp.unit
            ws.cell(r, COL_QTY).value = comp_qty_formula
            # 0.# = «5» для целых, «5.5» для дробных — единый стиль колонки.
            ws.cell(r, COL_QTY).number_format = "#,##0.#"
            ws.cell(r, COL_FORMULA).value = audit_text
            r += 1
            subnum += 1
        r += 1   # blank separator between groups
    return r
