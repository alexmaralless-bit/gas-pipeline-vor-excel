"""
interpret.py — turn raw dwg extract (io_schema.ExtractDocument) into
engineering-meaningful objects for the earthwork engine.

Responsibilities:
  * Parse trace-layer names like ``!!ГСН3_Г1_ГП@110``,
    ``!!ГСН3_Г1_ГП_Вводы@63`` and ``!ГСН_Г1_ПЭ_@110x10``
    → (stage, pressure_category,
    branch_role, dn_mm).
  * Parse numeric DN/DN2 from the (textual) ``DN`` attribute of a block,
    e.g. "Седловой отвод DN110х63" → (110, 63).
  * Group polylines by (branch_role, pressure_category, dn) → Branch +
    PipeRow list.
  * Provide simple counters for fitting-blocks (e.g. count of СО_КТ_*).

CallSite chooses ``Section`` (soil/season/laying_method) externally —
typically via CLI flags — and supplies it as the ``default_section``.
"""
from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from typing import Iterable, Optional

from .io_schema import BlockRecord, ExtractDocument, PolylineRecord
from .trench_profile import Branch, PipeRow, Section


# ============================================================================
# Layer-name parser
# ============================================================================

# !!ГСН<stage>_Г<pressure>_ГП[_<branch_type>]@<DN>
# Examples:
#   !!ГСН3_Г1_ГП@110          → stage=3 pressure=1 type=None DN=110
#   !!ГСН3_Г1_ГП_Вводы@63     → stage=3 pressure=1 type="Вводы" DN=63
#   !!ГСН2_Г3_ГП@90           → stage=2 pressure=3 type=None DN=90
TRACE_LAYER_LEGACY_RE = re.compile(
    r"^!{1,2}ГСН(?P<stage>\d+)_Г(?P<pressure>\d+)_ГП"
    r"(?:_(?P<branch_type>[^@]+))?"
    r"@(?P<dn>\d+)$"
)

# Newer project drawings use PE material layers without a stage digit and
# include wall thickness after DN:
#   !ГСН_Г1_ПЭ_@110x10       → stage=1 pressure=1 type=None DN=110
#   !ГСН_Г1_ПЭ_Вводы@63      → stage=1 pressure=1 type="Вводы" DN=63
TRACE_LAYER_PE_RE = re.compile(
    r"^!{1,2}ГСН(?P<stage>\d*)_Г(?P<pressure>\d+)_(?:ПЭ|ГП)"
    r"(?:_(?P<branch_type>[^@]*))?"
    r"@(?P<dn>\d+)(?:[xхXХ]\d+(?:[.,]\d+)?)?$"
)


@dataclass(frozen=True)
class TraceLayerKey:
    """Parsed components of a trace-layer name."""
    stage: int
    pressure_category: str          # "Г1", "Г3", ...
    branch_type: Optional[str]      # e.g. "Вводы"; None for main
    dn_mm: int

    @property
    def branch_role(self) -> str:
        """Map branch_type → human role."""
        if not self.branch_type:
            return "main"
        bt = self.branch_type.lower()
        if "ввод" in bt:
            return "consumer_inlets"
        return bt


def parse_trace_layer(layer_name: str) -> Optional[TraceLayerKey]:
    """Parse known gas trace-layer names. Returns None if not a trace layer."""
    m = TRACE_LAYER_LEGACY_RE.match(layer_name) or TRACE_LAYER_PE_RE.match(layer_name)
    if not m:
        return None
    stage_raw = m.group("stage")
    branch_type = m.group("branch_type")
    if branch_type is not None:
        branch_type = branch_type.strip("_") or None
    return TraceLayerKey(
        stage=int(stage_raw) if stage_raw else 1,
        pressure_category=f"Г{m.group('pressure')}",
        branch_type=branch_type,
        dn_mm=int(m.group("dn")),
    )


# ============================================================================
# Block-attribute DN parser
# ============================================================================

# Pattern variants found in the reference project:
#   "DN110"                          → 110
#   "Седловой отвод DN110х63"        → 110, 63    (cyrillic х separator)
#   "Седловой отвод DN110x63"        → 110, 63    (latin x)
#   "Переход DN315x225"              → 315, 225
#   "Отвод 90%%d с ЗН DN110"         → 110
#   "Кран КШ-150с в ограждении"      → no DN match → None
#   "Ду300"                          → 300  (alt notation)
_DN_RE = re.compile(r"[DД][NнУуYy]\s*(\d+)(?:\s*[xхXХ]\s*(\d+))?", re.IGNORECASE)


def parse_dn_from_attr(text: str) -> tuple[Optional[int], Optional[int]]:
    """Parse DN (primary) and DN2 (secondary, for transitions) from text.

    Returns (dn_mm, dn2_mm). Either or both may be None.
    """
    if not text:
        return (None, None)
    m = _DN_RE.search(text)
    if not m:
        return (None, None)
    dn = int(m.group(1))
    dn2 = int(m.group(2)) if m.group(2) else None
    return (dn, dn2)


# ============================================================================
# Geometric host_dn resolver
# ============================================================================

def _point_to_segment_distance_xy(
    px: float, py: float,
    ax: float, ay: float,
    bx: float, by: float,
) -> float:
    """Distance from point P to line segment AB in 2D (XY plane).

    Z is ignored — газопроводы и КТ/ЗУ-кран на чертеже планово близки,
    а перепад по Z отражает уровень земли, не близость trace-pipeline.
    """
    dx, dy = bx - ax, by - ay
    seg_len_sq = dx * dx + dy * dy
    if seg_len_sq < 1e-12:
        # Degenerate segment → distance to A
        ddx, ddy = px - ax, py - ay
        return (ddx * ddx + ddy * ddy) ** 0.5
    # Project P onto AB, clamp t to [0, 1]
    t = ((px - ax) * dx + (py - ay) * dy) / seg_len_sq
    t = max(0.0, min(1.0, t))
    cx, cy = ax + t * dx, ay + t * dy
    ddx, ddy = px - cx, py - cy
    return (ddx * ddx + ddy * ddy) ** 0.5


def resolve_geometric_host_dn(
    block_position: tuple[float, float, float],
    polylines: Iterable[PolylineRecord],
    *,
    max_distance: float = 10.0,
) -> Optional[int]:
    """Find DN основного газопровода под/около block.position через геометрию.

    Используется для блоков без DN-атрибута (КТ, ЗУ-кран): берётся
    ближайшая trace-полилиния (та, чьё имя слоя матчит ``parse_trace_layer``),
    из её layer name парсится DN.

    ``max_distance`` отсекает блоки которые "слишком далеко" от любой trace
    (например стоит в легенде чертежа, а не на трассе). Если ни одна
    polyline не ближе, возвращает ``None``.

    Distance считается в drawing units (метры для эталона). 10 м — разумный
    порог для масштаба 1:500-1:2000.
    """
    px, py, _ = block_position
    best_dn: Optional[int] = None
    best_dist = max_distance
    for p in polylines:
        if p.is_closed:
            continue
        key = parse_trace_layer(p.layer)
        if key is None:
            continue   # not a trace polyline
        ax, ay, _ = p.start
        bx, by, _ = p.end
        d = _point_to_segment_distance_xy(px, py, ax, ay, bx, by)
        if d < best_dist:
            best_dist = d
            best_dn = key.dn_mm
    return best_dn


# ============================================================================
# Aggregation
# ============================================================================

@dataclass
class TraceAggregate:
    """Sum of polyline lengths for one (branch_role, pressure, dn) bucket."""
    key: TraceLayerKey
    polyline_count: int
    total_length: float
    source_layers: list[str]


def aggregate_traces(polylines: Iterable[PolylineRecord]) -> list[TraceAggregate]:
    """Group polylines by parsed trace-layer key, sum lengths.

    Polylines on layers that don't match the trace pattern are ignored
    (they're scenery, dimensions, annotations, etc.).
    """
    buckets: dict[tuple, dict] = {}
    for p in polylines:
        if p.is_closed:
            continue
        key = parse_trace_layer(p.layer)
        if key is None:
            continue
        bucket_key = (key.branch_role, key.pressure_category, key.dn_mm)
        bucket = buckets.setdefault(bucket_key, {
            "key": key,
            "polyline_count": 0,
            "total_length": 0.0,
            "source_layers": set(),
        })
        bucket["polyline_count"] += 1
        bucket["total_length"] += p.length
        bucket["source_layers"].add(p.layer)
    return [
        TraceAggregate(
            key=b["key"],
            polyline_count=b["polyline_count"],
            total_length=b["total_length"],
            source_layers=sorted(b["source_layers"]),
        )
        for b in buckets.values()
    ]


# ============================================================================
# Branch builder
# ============================================================================

def build_branches(extract: ExtractDocument, default_section: Section) -> list[Branch]:
    """Build pipe_engine.Branch list from a raw dwg extract.

    Branches are formed by (branch_role, pressure_category, stage). For each,
    the PipeRow list is one entry per unique DN in that group.

    All branches get ``default_section`` — caller can override per-branch
    later if needed (e.g. via CLI flags or a separate config layer).
    """
    aggs = aggregate_traces(extract.polylines)
    # Group by (branch_role, pressure_category, stage)
    grouped: dict[tuple, list[TraceAggregate]] = defaultdict(list)
    for a in aggs:
        grouped[(a.key.branch_role, a.key.pressure_category, a.key.stage)].append(a)

    branches: list[Branch] = []
    for (role, pressure, stage), items in sorted(grouped.items()):
        # Sort pipe rows by DN for deterministic output
        items.sort(key=lambda x: x.key.dn_mm)
        pipes = [
            PipeRow(dn_mm=item.key.dn_mm, length_grp1=item.total_length, length_other=0.0,
                    gnb_length=0.0)
            for item in items
        ]
        branch_id = f"{role}_{pressure.lower()}_stage{stage}"
        branches.append(Branch(id=branch_id, section=default_section, pipes=pipes))
    return branches


# ============================================================================
# Fitting counters
# ============================================================================

@dataclass(frozen=True)
class FittingCount:
    block_name: str
    count: int


def count_fittings(blocks: Iterable[BlockRecord]) -> list[FittingCount]:
    """Count block instances by effective name."""
    counts: dict[str, int] = defaultdict(int)
    for b in blocks:
        counts[b.block_name_effective] += 1
    return [FittingCount(name, n) for name, n in sorted(counts.items())]


def count_by_name(blocks: Iterable[BlockRecord], name: str) -> int:
    """Count instances of a specific block by effective name."""
    return sum(1 for b in blocks if b.block_name_effective == name)


def count_consumer_inlet_polylines(extract: ExtractDocument) -> int:
    """Number of polylines on consumer-inlet trace layers (`branch_role="consumer_inlets"`).

    In Russian gas-distribution dwg practice each polyline on a layer like
    ``!!ГСН3_Г1_ГП_Вводы@63`` represents ONE service line from the main
    pipeline to ONE consumer's gate (Узел А). So polyline count = consumer
    count = number of «Узлы А и Б» exit-sand nodes.
    """
    return sum(
        a.polyline_count for a in aggregate_traces(extract.polylines)
        if a.key.branch_role == "consumer_inlets"
    )


# ============================================================================
# Exit-sand node derivation
# ============================================================================

# Default sand_units multipliers for the canonical exit-node types, per
# эталонной книги rows 30-36 (the reference project's «Песок для устройства
# выходов из земли» block).
EXIT_SAND_DEFAULTS = {
    "ЗУ надземно":      4,   # ground-level shut-off valve (one pad at each of 4 corners)
    "Узел А":           2,   # consumer connection point — one pad + one КТ
    "КТ по трассе":     1,   # standalone control tube
    "Пересеч инж комм": 1,   # crossing of utility line in casing
    "ГНБ переходы":     1,   # exits at HDD entry/exit pits
}


def derive_exit_sand_from_extract(
    extract: ExtractDocument,
    *,
    hand_trench_count: int = 0,
    gnb_pit_count: int = 0,
    joint_kt_count: int = 0,
) -> list:
    """Build the canonical exit_sand node list from an ExtractDocument.

    Mapping rules (derived from эталонной книги + scout of real dwg):

    * **ЗУ надземно (и КТ)** ← ``count(СО_ЗУ_кран)``, sand_units=4.
      The above-ground shut-off valve sits on a 4-corner sand pad.
    * **Узел А** ← ``count_consumer_inlet_polylines(extract)``, sand_units=2.
      One ввод polyline = one consumer service line = one Узел А (+КТ).
    * **КТ по трассе** ← ``count(СО_КТ_по_трассе)``, sand_units=1.
      Standalone control tube above straight pipeline.
    * **Пересеч инж комм в футляре** ← ``hand_trench_count`` (CLI param),
      sand_units=1. Each crossing of a utility line is dug by hand AND
      gets a sand bed under the casing.
    * **Прокладка в футляре ГНБ** ← ``gnb_pit_count`` (CLI param),
      sand_units=1. Each HDD section has sand pads at its working/receiving
      pits.
    * **КТ при совм. прокладке** ← ``joint_kt_count`` (CLI param),
      sand_units=1. Rare.

    Blocks **NOT** in the list (intentionally):

    * **СО_Загл_ПЭ (заглушка)** — just a pipe end-cap below ground, no
      surface exit, no sand pad needed.
    * **СО_Седло_Отвод, СО_Тройник, СО_Переход, СО_Муфта_с_ЗН, СО_Отвод**
      — all internal pipe fittings, fully buried, no surface exit.

    Returns a list of ExitSandNode (lazy import to avoid circular dep).
    """
    from .earthwork import ExitSandNode  # local import — avoids module cycle

    fitting_counts = {c.block_name: c.count for c in count_fittings(extract.blocks)}

    nodes = []

    zu_count = fitting_counts.get("СО_ЗУ_кран", 0)
    if zu_count > 0:
        nodes.append(ExitSandNode("ЗУ надземно (и КТ)", zu_count, EXIT_SAND_DEFAULTS["ЗУ надземно"]))

    inlet_polys = count_consumer_inlet_polylines(extract)
    if inlet_polys > 0:
        nodes.append(ExitSandNode("Узел А (вводы)", inlet_polys, EXIT_SAND_DEFAULTS["Узел А"]))

    kt_count = fitting_counts.get("СО_КТ_по_трассе", 0)
    if kt_count > 0:
        nodes.append(ExitSandNode("КТ по трассе", kt_count, EXIT_SAND_DEFAULTS["КТ по трассе"]))

    if joint_kt_count > 0:
        nodes.append(ExitSandNode("КТ при совм. прокладке", joint_kt_count, 1))

    if hand_trench_count > 0:
        nodes.append(ExitSandNode("Пересеч инж комм в футляре", hand_trench_count, EXIT_SAND_DEFAULTS["Пересеч инж комм"]))

    if gnb_pit_count > 0:
        nodes.append(ExitSandNode("Прокладка в футляре ГНБ", gnb_pit_count, EXIT_SAND_DEFAULTS["ГНБ переходы"]))

    return nodes
