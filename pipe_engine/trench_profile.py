"""
trench_profile.py — справочники нормативных значений и dataclass'ы для
профиля траншеи газораспределительной сети.

Нормативные источники (см. ``norms-agent``):

* СНиП 12-04-2002, Приложение 1, Таблица 1 — допустимая крутизна откоса
  незакреплённой выемки в зависимости от грунта и глубины.
* СП 45.13330.2017, п. 6.1.3 + Таблица 6.1 — минимальная ширина траншеи
  для трубопроводов кроме магистральных (= распределительные газовые сети).
* СП 45.13330.2017, п. 6.1.15 — в мёрзлых грунтах допускается
  увеличение вертикальных стенок на величину промерзания (зимний режим
  → ``slope_ratio=0`` нормативно).
* СП 62.13330.2011 — газораспределительные системы (ширина траншеи
  отсылает к СП 45.13330; задаёт глубину заложения, подушку над верхом).

Все линейные размеры в метрах. DN везде в миллиметрах если не указано.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Literal, Optional


# ============================================================================
# Грунты + Уклоны откосов (СНиП 12-04-2002, Прил.1, Табл.1)
# ============================================================================

class SoilType(str, Enum):
    HEAPED_LOOSE = "насыпной_неслежавшийся"
    SAND         = "песчаный"
    SANDY_LOAM   = "супесь"
    LOAM         = "суглинок"
    CLAY         = "глина"
    LOESS        = "лёссовый"


# СНиП 12-04-2002 Прил.1 Табл.1 — (max_depth_m, slope_ratio) ascending по глубине.
# slope_ratio = k где откос записывается как 1:k (k=0 → вертикально).
SLOPE_TABLE_SNIP_12_04_2002: dict[SoilType, list[tuple[float, float]]] = {
    SoilType.HEAPED_LOOSE: [(1.5, 0.67), (3.0, 1.00), (5.0, 1.25)],
    SoilType.SAND:         [(1.5, 0.50), (3.0, 1.00), (5.0, 1.00)],
    SoilType.SANDY_LOAM:   [(1.5, 0.25), (3.0, 0.67), (5.0, 0.85)],
    SoilType.LOAM:         [(1.5, 0.00), (3.0, 0.50), (5.0, 0.75)],
    SoilType.CLAY:         [(1.5, 0.00), (3.0, 0.25), (5.0, 0.50)],
    SoilType.LOESS:        [(1.5, 0.00), (3.0, 0.50), (5.0, 0.50)],
}


def slope_by_snip(soil: SoilType, total_depth_m: float) -> float:
    """Допустимый откос по СНиП 12-04-2002 Прил.1 Табл.1.

    Возвращает k где откос пишется как 1:k.
    Для глубин > 5 м берётся последняя строка (требует расчёта по СП 45.13330
    Прил.В, что в этом порте не реализовано — поднимать предупреждение).
    """
    table = SLOPE_TABLE_SNIP_12_04_2002[soil]
    for max_h, k in table:
        if total_depth_m <= max_h:
            return k
    return table[-1][1]


# ============================================================================
# Ширина дна траншеи (СП 45.13330.2017 п.6.1.3 + Табл.6.1)
# ============================================================================

LayingMethod = Literal["plet", "individual"]   # плеть / отдельные трубы


def bottom_width_by_sp45(
    diameter_outer_m: float,
    *,
    laying_method: LayingMethod = "plet",
    min_width: float = 0.7,
) -> float:
    """Минимальная ширина дна траншеи по СП 45.13330.2017 Табл.6.1.

    Для трубопроводов кроме магистральных (= газораспределительные сети).
    Для сварного стыка:
      * плетями, D ≤ 0.7 м:  D + 0.3, но не менее 0.7 м (=``min_width``);
      * плетями, D > 0.7 м:  D + 0.2;
      * отдельными трубами, D ≤ 0.5 м:  D + 0.5;
      * отдельными трубами, 0.5 < D ≤ 1.6 м:  D + 0.8.

    На практике для всех ПЭ DN ≤ 400 мм укладкой плетями результат = 0.7 м.
    """
    if laying_method == "plet":
        add = 0.3 if diameter_outer_m <= 0.7 else 0.2
    elif laying_method == "individual":
        if diameter_outer_m <= 0.5:
            add = 0.5
        elif diameter_outer_m <= 1.6:
            add = 0.8
        else:
            add = 1.4
    else:
        raise ValueError(f"unknown laying_method: {laying_method!r}")
    return max(diameter_outer_m + add, min_width)


# ============================================================================
# Профиль траншеи + Section + Branch
# ============================================================================

Season = Literal["winter", "summer"]


@dataclass(frozen=True)
class TrenchProfile:
    """Поперечный профиль траншеи на одном участке.

    Геометрия — трапеция:
        верх = bottom_width + 2 × total_depth × slope_ratio
        площадь = (bottom_width + total_depth × slope_ratio) × total_depth
    """
    bottom_width: float       # м, низ траншеи
    slope_ratio: float = 0.0  # 1:k, 0 = вертикально

    def cross_section_area(self, total_depth_m: float) -> float:
        """Площадь поперечного сечения, м²."""
        return (self.bottom_width + total_depth_m * self.slope_ratio) * total_depth_m

    def top_width(self, total_depth_m: float) -> float:
        """Ширина траншеи поверху, м."""
        return self.bottom_width + 2.0 * total_depth_m * self.slope_ratio

    def soil_volume_per_meter(self, total_depth_m: float) -> float:
        """Объём грунта на 1 п.м. трассы, м³."""
        return self.cross_section_area(total_depth_m)


@dataclass
class Section:
    """Участок трассы со своими условиями (грунт, сезон, способ укладки).

    Привязка к трубам — через ``Branch.section``. Один Section может
    применяться к нескольким веткам.

    Если ``forced_slope_ratio`` или ``forced_bottom_width`` заданы,
    они перебивают нормативный автоподбор по СНиП/СП — это позволяет
    инженеру отступать от норматива по проектному решению.
    """
    name: str                                # отображаемое имя, для логов/отчётов
    soil: SoilType = SoilType.SAND
    season: Season = "summer"
    laying_method: LayingMethod = "plet"
    sand_cushion_required: bool = False      # подушка снизу + присыпка сверху из ПРИВОЗНОГО песка
    sand_side_backfill_required: bool = False  # пазухи трубы засыпать ПРИВОЗНЫМ песком
    forced_slope_ratio: Optional[float] = None
    forced_bottom_width: Optional[float] = None

    def trench_profile(self, diameter_outer_m: float, total_depth_m: float) -> TrenchProfile:
        """Подобрать профиль для трубы данного диаметра на этом участке."""
        # 1) Уклон откоса
        if self.forced_slope_ratio is not None:
            slope = self.forced_slope_ratio
        elif self.season == "winter":
            # СП 45.13330.2017 п.6.1.15: мёрзлый грунт держит вертикальную стенку.
            slope = 0.0
        else:
            slope = slope_by_snip(self.soil, total_depth_m)

        # 2) Ширина дна
        if self.forced_bottom_width is not None:
            bottom = self.forced_bottom_width
        else:
            bottom = bottom_width_by_sp45(
                diameter_outer_m, laying_method=self.laying_method
            )
        return TrenchProfile(bottom_width=bottom, slope_ratio=slope)


@dataclass
class PipeRow:
    """Один участок трубы фиксированного DN."""
    dn_mm: int
    length_grp1: float = 0.0   # длина в зоне ГРП-1, м
    length_other: float = 0.0  # длина по прочим участкам, м
    gnb_length: float = 0.0    # часть длины протянутая ГНБ (вычитается из открытой)


@dataclass
class Branch:
    """Ветка газопровода — связный сегмент трассы со своим участком.

    Одна ветка может содержать несколько PipeRow разных DN (например,
    переход с DN 160 на DN 110 далее по трассе).
    """
    id: str
    section: Section
    pipes: list[PipeRow]
    joint_with: Optional[str] = None   # id другой ветки, если совместная прокладка


@dataclass
class JointPipePair:
    """Один участок совместной прокладки двух труб в одной траншее.

    В реальности это две параллельные трубы лежащие в одной более широкой
    траншее (см. ИсхДанные!$B$11 = joint_bottom_width, обычно 1.2 м).
    Геометрия: трапеция как обычно, но с увеличенной шириной дна.
    Объём грунта считается по большему из двух DN (max(dn1, dn2)) и
    суммарному объёму двух труб J = π × ((D1/1000)² + (D2/1000)²) × L / 4.
    """
    id: str
    dn1_mm: int                # один из DN
    dn2_mm: int                # второй DN
    length_m: float            # длина совместного участка, м
