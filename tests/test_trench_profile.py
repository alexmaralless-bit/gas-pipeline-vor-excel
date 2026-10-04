"""
Unit tests for pipe_engine.trench_profile — справочники СНиП/СП и
геометрия трапеции.
"""
from __future__ import annotations

import math
import unittest

from pipe_engine.trench_profile import (
    Branch,
    PipeRow,
    Section,
    SoilType,
    TrenchProfile,
    bottom_width_by_sp45,
    slope_by_snip,
)


class TestSnipSlope(unittest.TestCase):
    """СНиП 12-04-2002 Прил.1 Табл.1 — крутизна откоса."""

    def test_sand_at_15m(self):
        # песчаный грунт, глубина 1.5 м → 1:0.5
        self.assertEqual(slope_by_snip(SoilType.SAND, 1.5), 0.5)

    def test_sand_at_05m(self):
        # неглубокая траншея — берётся первая строка (h ≤ 1.5)
        self.assertEqual(slope_by_snip(SoilType.SAND, 0.5), 0.5)

    def test_sand_at_30m(self):
        # глубокая — следующая строка
        self.assertEqual(slope_by_snip(SoilType.SAND, 3.0), 1.0)

    def test_clay_at_15m_vertical(self):
        # глина при h ≤ 1.5 м — вертикально
        self.assertEqual(slope_by_snip(SoilType.CLAY, 1.5), 0.0)

    def test_clay_at_30m(self):
        self.assertEqual(slope_by_snip(SoilType.CLAY, 3.0), 0.25)

    def test_loam_15m_vertical(self):
        # суглинок при h ≤ 1.5 м — вертикально
        self.assertEqual(slope_by_snip(SoilType.LOAM, 1.5), 0.0)

    def test_sandy_loam_15m(self):
        self.assertEqual(slope_by_snip(SoilType.SANDY_LOAM, 1.5), 0.25)

    def test_heaped_15m(self):
        self.assertEqual(slope_by_snip(SoilType.HEAPED_LOOSE, 1.5), 0.67)

    def test_loess_30m(self):
        self.assertEqual(slope_by_snip(SoilType.LOESS, 3.0), 0.5)

    def test_over_5m_falls_to_last(self):
        # h > 5 → используется значение для h=5 (требует расчёта по СП 45 Прил.В)
        self.assertEqual(slope_by_snip(SoilType.SAND, 7.0), 1.0)


class TestSp45BottomWidth(unittest.TestCase):
    """СП 45.13330.2017 п.6.1.3 + Табл.6.1 — ширина дна траншеи."""

    def test_pe63_plet(self):
        # DN 63 (D=0.063 м), плеть → D+0.3 = 0.363 → min 0.7
        self.assertEqual(bottom_width_by_sp45(0.063), 0.7)

    def test_pe225_plet(self):
        # DN 225 (D=0.225 м), плеть → D+0.3 = 0.525 → min 0.7
        self.assertEqual(bottom_width_by_sp45(0.225), 0.7)

    def test_pe400_plet(self):
        # DN 400 (D=0.4 м), плеть → D+0.3 = 0.7 (точно min)
        self.assertAlmostEqual(bottom_width_by_sp45(0.400), 0.7, places=9)

    def test_pe630_plet(self):
        # DN 630 (D=0.63), плеть, D <= 0.7 → D+0.3 = 0.93
        self.assertAlmostEqual(bottom_width_by_sp45(0.630), 0.93, places=9)

    def test_pe800_plet(self):
        # D > 0.7 м, плеть → D + 0.2
        self.assertAlmostEqual(bottom_width_by_sp45(0.800), 1.00, places=9)

    def test_steel_individual_small(self):
        # отдельными трубами, D ≤ 0.5: D + 0.5
        self.assertAlmostEqual(bottom_width_by_sp45(0.219, laying_method="individual"), 0.719, places=9)

    def test_steel_individual_med(self):
        # отдельными трубами, 0.5 < D ≤ 1.6: D + 0.8
        self.assertAlmostEqual(bottom_width_by_sp45(1.0, laying_method="individual"), 1.8, places=9)


class TestTrenchProfileGeometry(unittest.TestCase):
    """Геометрия трапеции."""

    def test_rectangle_when_vertical(self):
        # slope=0 — площадь = bottom × depth (прямоугольник)
        p = TrenchProfile(bottom_width=0.7, slope_ratio=0.0)
        self.assertAlmostEqual(p.cross_section_area(1.5), 0.7 * 1.5, places=9)
        self.assertEqual(p.top_width(1.5), 0.7)

    def test_picture_sample(self):
        # Картинка из обсуждения: depth=1.5 м, bottom=0.7, slope=0.5 → top=2.2 м
        # Площадь = (0.7 + 2.2)/2 × 1.5 = 2.175 м²
        p = TrenchProfile(bottom_width=0.7, slope_ratio=0.5)
        self.assertAlmostEqual(p.top_width(1.5), 2.2, places=9)
        self.assertAlmostEqual(p.cross_section_area(1.5), 2.175, places=9)

    def test_slope_doubles_volume(self):
        # Заявленное в обсуждении: откос 1:0.5 при 1.5 м увеличивает объём ~2x
        rect = TrenchProfile(bottom_width=0.7).cross_section_area(1.5)
        trap = TrenchProfile(bottom_width=0.7, slope_ratio=0.5).cross_section_area(1.5)
        self.assertAlmostEqual(trap / rect, 2.175 / 1.05, places=5)


class TestSectionProfilePicker(unittest.TestCase):
    """Логика Section.trench_profile — норматив + перебивки."""

    def test_winter_overrides_to_vertical(self):
        # Зимой stope=0 невзирая на soil/depth (СП 45.13330 п.6.1.15)
        s = Section(name="t", soil=SoilType.SAND, season="winter")
        p = s.trench_profile(diameter_outer_m=0.110, total_depth_m=2.0)
        self.assertEqual(p.slope_ratio, 0.0)
        self.assertEqual(p.bottom_width, 0.7)  # DN 110 плетью

    def test_summer_sand_15m(self):
        # Летом песок при 1.5 м → 1:0.5
        s = Section(name="t", soil=SoilType.SAND, season="summer")
        p = s.trench_profile(diameter_outer_m=0.160, total_depth_m=1.5)
        self.assertEqual(p.slope_ratio, 0.5)

    def test_summer_loam_15m(self):
        # Суглинок летом при 1.5 м → вертикально (по таблице)
        s = Section(name="t", soil=SoilType.LOAM, season="summer")
        p = s.trench_profile(diameter_outer_m=0.160, total_depth_m=1.5)
        self.assertEqual(p.slope_ratio, 0.0)

    def test_forced_slope_overrides_norms(self):
        # forced_slope перебивает норматив
        s = Section(name="t", soil=SoilType.SAND, season="summer",
                    forced_slope_ratio=0.0)
        p = s.trench_profile(diameter_outer_m=0.063, total_depth_m=1.5)
        self.assertEqual(p.slope_ratio, 0.0)

    def test_forced_bottom_width_overrides(self):
        s = Section(name="t", soil=SoilType.SAND, forced_bottom_width=1.5)
        p = s.trench_profile(diameter_outer_m=0.063, total_depth_m=1.5)
        self.assertEqual(p.bottom_width, 1.5)


class TestBranchDataclass(unittest.TestCase):
    """Простая sanity-check вложенных dataclass'ов."""

    def test_branch_contains_pipes(self):
        s = Section(name="main", soil=SoilType.SAND, season="winter")
        b = Branch(id="b1", section=s, pipes=[
            PipeRow(dn_mm=110, length_grp1=100.0),
            PipeRow(dn_mm=63, length_grp1=200.0),
        ])
        self.assertEqual(len(b.pipes), 2)
        self.assertEqual(b.section.name, "main")
        self.assertIsNone(b.joint_with)


if __name__ == "__main__":
    unittest.main()
