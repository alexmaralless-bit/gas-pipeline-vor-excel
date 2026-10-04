"""
earthwork.py — расчёт земляных работ для газораспределительной сети.

Чистая порт-функция исходного листа «Расчет земли» из эталонной книги,
переработанная под:

  * параметрический профиль траншеи (откос + ширина дна по СНиП/СП —
    см. ``pipe_engine.trench_profile``);
  * привязку к участкам через ``Branch.section`` (разные части трассы
    могут иметь разный грунт, сезон, способ укладки);
  * сезонность (снег считается только при season="winter").

Регрессия Эталон воспроизводится приблизительно (см. tests/test_earthwork.py):
несколько ячеек дрейфуют на 2-5% потому что мы лечим известные баги
исходного xlsm — формула ширины траншеи теперь единообразна для всех DN,
а sand_pad учитывается только если песчаная подушка явно требуется.

Все объёмы возвращаются в м³, длины в м, DN в мм если не указано иное.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Optional

from .trench_profile import Branch, PipeRow, Section, TrenchProfile


# ============================================================================
# Project-wide inputs
# ============================================================================

@dataclass
class ProjectParams:
    """Проектные параметры верхнего уровня (ИсхДанные)."""
    depth: float                       # B4 — глубина заложения, м (до верха трубы)
    snow_thickness: float = 0.5        # ИсхДанные!D2 — толщина снежного покрова, м
    sand_pad: float = 0.1              # B5 — подушка / присыпка из песка, м
    snow_corridor_width: float = 3.2   # C53 — ширина коридора расчистки от снега
    excavator_loosening: float = 1.05  # C57 — коэффициент разрыхления, доли ед.
    gnb_exit_sand_size: float = 0.35   # B7 — сторона песчаной "коробки" вокруг выхода трубы, м

    @property
    def sand_per_gnb_exit(self) -> float:
        """B7 = 0.35 × 0.35 × depth — объём песка на 1 выход трубы из земли, м³."""
        return self.gnb_exit_sand_size ** 2 * self.depth


# ============================================================================
# Прочие входы — узлы выходов из земли и котлованы ГНБ
# ============================================================================

@dataclass
class ExitSandNode:
    """Один тип узла в блоке «Песок для устройства выходов из земли»."""
    name: str                # отображаемое имя
    count: int               # количество узлов этого типа
    sand_units: float        # множитель B7 на 1 узел (4 / 2 / 1 типично)


@dataclass
class PitParams:
    """Котлованы ГНБ + крепление щитами."""
    working_pit_count: int
    working_pit_vol: float = 8.0        # 2×2×2 м по умолчанию
    receiving_pit_count: int = 0
    receiving_pit_vol: float = 0.0
    shoring_perimeter: float = 8.0      # (2+2)*2 м периметр одного котлована
    shoring_height_addition: float = 0.2  # добавка к глубине для высоты щитов


# ============================================================================
# Геометрические помощники
# ============================================================================

def pipe_volume_m3(dn_mm: float, length_m: float) -> float:
    """Внутренний объём трубы: π × (D/1000)² / 4 × L (м³)."""
    return math.pi * (dn_mm / 1000.0) ** 2 / 4.0 * length_m


def total_trench_depth(depth_m: float, dn_mm: float, sand_pad_m: float,
                        *, sand_cushion: bool) -> float:
    """Полная глубина траншеи = заложение + диаметр трубы [+ песчаная подушка]."""
    return depth_m + dn_mm / 1000.0 + (sand_pad_m if sand_cushion else 0.0)


# ============================================================================
# Расчёт по веткам
# ============================================================================

@dataclass
class PipeRowResult:
    """Результат расчёта одной строки PipeRow (одна ветка × один DN)."""
    branch_id: str
    section_name: str
    dn_mm: int

    total_length: float      # G — общая длина (грп1 + прочее)
    gnb_length: float        # length cut by ГНБ
    open_length: float       # I — длина открытой траншеи
    trench_depth: float      # глубина траншеи, м
    profile: TrenchProfile   # геометрия

    pipe_volume: float       # J — объём самих труб, м³
    soil_volume: float       # K — объём грунта (выемка), м³
    sand_side_volume: float  # L — песок по сторонам трубы, м³
    sand_cushion_volume: float  # M — песок подушки/присыпки, м³


def _row_result(branch: Branch, row: PipeRow, params: ProjectParams) -> PipeRowResult:
    sec = branch.section
    dn_m = row.dn_mm / 1000.0

    total = row.length_grp1 + row.length_other
    open_len = total - row.gnb_length
    if open_len < 0:
        open_len = 0.0

    depth_total = total_trench_depth(
        params.depth, row.dn_mm, params.sand_pad,
        sand_cushion=sec.sand_cushion_required,
    )
    profile = sec.trench_profile(dn_m, depth_total)

    j_pipe = pipe_volume_m3(row.dn_mm, open_len)
    k_soil = profile.cross_section_area(depth_total) * open_len

    # L (sand at sides): объём песчаной засыпки пазух пробоя.
    # Геометрически: площадь поперечного сечения «коробки» с шириной trench.bottom
    # на высоте трубы (= D), минус сама труба = пазухи. ×1 если требуется.
    if sec.sand_side_backfill_required:
        l_sand_side = (dn_m * profile.bottom_width * open_len) - j_pipe
        if l_sand_side < 0:
            l_sand_side = 0.0
    else:
        l_sand_side = 0.0

    # M (sand cushion+sprinkle): подушка снизу + присыпка сверху = 10 см × ширина × длина.
    if sec.sand_cushion_required:
        m_sand_cushion = open_len * profile.bottom_width * params.sand_pad
    else:
        m_sand_cushion = 0.0

    return PipeRowResult(
        branch_id=branch.id,
        section_name=sec.name,
        dn_mm=row.dn_mm,
        total_length=total,
        gnb_length=row.gnb_length,
        open_length=open_len,
        trench_depth=depth_total,
        profile=profile,
        pipe_volume=j_pipe,
        soil_volume=k_soil,
        sand_side_volume=l_sand_side,
        sand_cushion_volume=m_sand_cushion,
    )


@dataclass
class BranchAggregate:
    """Агрегированные показатели по проекту (сумма по всем веткам и DN)."""
    rows: list[PipeRowResult]
    snow_clearing_volume: float   # C53
    total_pipe_volume: float      # J ИТОГО
    total_soil_volume: float      # K ИТОГО
    total_sand_side: float        # L ИТОГО
    total_sand_cushion: float     # M ИТОГО
    total_open_length: float      # I ИТОГО

    # Под-агрегаты, нужные баланс-блоку (C53..F63):
    snow_open_length: float       # длина участков где season=winter (для C53)


def compute_branches(branches: list[Branch], params: ProjectParams) -> BranchAggregate:
    """Посчитать все ветки и собрать ИТОГО."""
    rows: list[PipeRowResult] = []
    snow_len = 0.0
    for br in branches:
        for pr in br.pipes:
            r = _row_result(br, pr, params)
            rows.append(r)
            if br.section.season == "winter":
                snow_len += r.open_length

    snow_vol = snow_len * params.snow_corridor_width * params.snow_thickness

    return BranchAggregate(
        rows=rows,
        snow_clearing_volume=snow_vol,
        total_pipe_volume=sum(r.pipe_volume for r in rows),
        total_soil_volume=sum(r.soil_volume for r in rows),
        total_sand_side=sum(r.sand_side_volume for r in rows),
        total_sand_cushion=sum(r.sand_cushion_volume for r in rows),
        total_open_length=sum(r.open_length for r in rows),
        snow_open_length=snow_len,
    )


# ============================================================================
# Песок на выходы из земли + Котлованы ГНБ
# ============================================================================

@dataclass
class ExitSandResult:
    nodes: list[dict]  # [{name, count, volume}, ...]
    total: float


def compute_exit_sand(nodes: list[ExitSandNode], params: ProjectParams) -> ExitSandResult:
    per = [
        {"name": n.name, "count": n.count,
         "volume": n.count * n.sand_units * params.sand_per_gnb_exit}
        for n in nodes
    ]
    return ExitSandResult(nodes=per, total=sum(x["volume"] for x in per))


@dataclass
class PitResult:
    working: float        # C46
    receiving: float      # C47
    excavation: float     # C48 = working + receiving
    backfill: float       # C49 = excavation
    shoring_area: float   # C50, м²


def compute_pits(pit: PitParams, params: ProjectParams) -> PitResult:
    working = pit.working_pit_count * pit.working_pit_vol
    receiving = pit.receiving_pit_count * pit.receiving_pit_vol
    excavation = working + receiving
    backfill = excavation
    shoring = pit.shoring_perimeter * (params.depth + pit.shoring_height_addition) * pit.working_pit_count
    return PitResult(working, receiving, excavation, backfill, shoring)


# ============================================================================
# Сводка земляных работ + Баланс
# ============================================================================

@dataclass
class EarthworksSummary:
    snow_clearing: float           # C53
    hand_trench_excavation: float  # C54
    hand_trench_backfill: float    # C55
    hand_trench_compact: float     # C56
    excavator_excavation: float    # C57
    sand_cushion_total: float      # C58
    sand_sprinkle_total: float     # C59
    landfill_volume: float         # C62 — излишек грунта в валик
    dozer_backfill: float          # C60
    dozer_compact: float           # C61
    excavation_total: float        # D63
    backfill_total: float          # E63
    exit_sand_recycled: float      # E64
    pipe_volume_recycled: float    # E65
    balance: float                 # F63 (должен быть ≈ 0)


def compute_earthworks_summary(
    aggregate: BranchAggregate,
    exit_sand: ExitSandResult,
    *,
    hand_trench_count: int,
    hand_trench_section_length: float,
    hand_trench_width: float,
    params: ProjectParams,
) -> EarthworksSummary:
    """Сводка земляных работ + проверка баланса масс.

    hand_trench_count       — число участков ручной траншеи (B35 в эталоне)
    hand_trench_section_length — длина одного участка ручной траншеи (4 м в эталоне)
    hand_trench_width       — ширина ручной траншеи (0.7 м в эталоне)
    """
    K_total = aggregate.total_soil_volume
    J_total = aggregate.total_pipe_volume
    L_total = aggregate.total_sand_side
    M_total = aggregate.total_sand_cushion

    c54 = hand_trench_count * hand_trench_section_length * hand_trench_width * params.depth
    c55 = c54
    c56 = c55
    c57 = (K_total - c54) * params.excavator_loosening
    c58 = M_total
    c59 = L_total + M_total * 2.0
    c62 = J_total + exit_sand.total + c58 + c59
    c60 = c57 - c62
    c61 = c60

    d63 = c54 + c57
    e63 = c55 + c58 + c59 + c60
    e64 = exit_sand.total
    e65 = J_total
    f63 = d63 - (e63 + e64 + e65)

    return EarthworksSummary(
        snow_clearing=aggregate.snow_clearing_volume,
        hand_trench_excavation=c54,
        hand_trench_backfill=c55,
        hand_trench_compact=c56,
        excavator_excavation=c57,
        sand_cushion_total=c58,
        sand_sprinkle_total=c59,
        landfill_volume=c62,
        dozer_backfill=c60,
        dozer_compact=c61,
        excavation_total=d63,
        backfill_total=e63,
        exit_sand_recycled=e64,
        pipe_volume_recycled=e65,
        balance=f63,
    )
