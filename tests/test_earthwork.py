"""
Regression test for pipe_engine.earthwork — refactored API.

Эталон воспроизводится как один Branch с зимним песчаным Section
(slope=0, как было в исходном xlsm). Несколько значений K (rows 12-14)
дрейфуют на 2-5% относительно xlsm потому что мы лечим известный баг
исходной формулы — sand_pad больше не добавляется к глубине когда
песчаная подушка не требуется. Баланс земляных масс продолжает сходиться.

Полный сравнительный дамп xlsm-фикстуры доступен в tests/fixtures/
earthwork_reference.json — для сверки можно открыть его руками.
"""
from __future__ import annotations

import json
import math
import unittest
from pathlib import Path

from pipe_engine.earthwork import (
    ExitSandNode,
    PitParams,
    ProjectParams,
    compute_branches,
    compute_earthworks_summary,
    compute_exit_sand,
    compute_pits,
)
from pipe_engine.trench_profile import (
    Branch,
    PipeRow,
    Section,
    SoilType,
)


FIXTURE_PATH = Path(__file__).resolve().parent / "fixtures" / "earthwork_reference.json"


def load_fixture() -> dict:
    with FIXTURE_PATH.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    return {rec["addr"]: rec for rec in data["records"]}


# ============================================================================
# Воссоздание эталона через новый API
# ============================================================================

def reference_project() -> ProjectParams:
    # B4=1.8, ИсхДанные!D2=0.5, B5=0.1
    return ProjectParams(depth=1.8, snow_thickness=0.5, sand_pad=0.1)


def reference_section() -> Section:
    """Один участок на весь проект — зимний песчаный (slope=0)."""
    return Section(
        name="Демо main",
        soil=SoilType.SAND,
        season="winter",                # → slope=0 по СП 45.13330 п.6.1.15
        laying_method="plet",           # → bottom_width = 0.7 для всех DN ≤ 400
        sand_cushion_required=False,    # местный грунт пригоден
        sand_side_backfill_required=False,
    )


def reference_branch() -> Branch:
    """Все 5 строк xlsm — одна ветка."""
    s = reference_section()
    return Branch(id="main", section=s, pipes=[
        PipeRow(dn_mm=63,  length_grp1=968.9,  length_other=587.0,  gnb_length=237.5),
        PipeRow(dn_mm=90,  length_grp1=1277.9, length_other=1004.0, gnb_length=153.6),
        PipeRow(dn_mm=225, length_grp1=5.0,    length_other=0.0,    gnb_length=0.0),
        PipeRow(dn_mm=160, length_grp1=836.6,  length_other=1877.0, gnb_length=603.2),
        PipeRow(dn_mm=110, length_grp1=684.9,  length_other=241.0,  gnb_length=263.2),
    ])


def reference_exit_nodes() -> list[ExitSandNode]:
    return [
        ExitSandNode("ЗУ надземно (и КТ)",                count=21,  sand_units=4),
        ExitSandNode("Узел А и Б (с КТ)",                  count=155, sand_units=2),
        ExitSandNode("КТ",                                 count=156, sand_units=1),
        ExitSandNode("КТ при совм прокладке",              count=0,   sand_units=1),
        ExitSandNode("Пересечение дороги/свеча/ГНБ",       count=0,   sand_units=1),
        ExitSandNode("Пересеч инж комм в футляре",         count=92,  sand_units=1),
        ExitSandNode("прокладка в футляре ГНБ",            count=44,  sand_units=1),
    ]


def reference_pit() -> PitParams:
    return PitParams(working_pit_count=72, working_pit_vol=8.0)


# ============================================================================
# Хелпер
# ============================================================================

def assert_close(testcase, actual, expected, label, rel_tol=1e-9, abs_tol=1e-9):
    if expected == 0.0:
        if abs(actual) > abs_tol:
            testcase.fail(f"{label}: expected 0, got {actual!r}")
        return
    if not math.isclose(actual, expected, rel_tol=rel_tol, abs_tol=abs_tol):
        testcase.fail(f"{label}: expected {expected!r}, got {actual!r} (rel diff {abs(actual-expected)/abs(expected):.2e})")


# ============================================================================
# Tests — values matching xlsm fixture (cell-for-cell)
# ============================================================================

class TestEarthworkExactXlsm(unittest.TestCase):
    """Эти ячейки совпадают с Excel'ом cell-for-cell."""

    @classmethod
    def setUpClass(cls):
        cls.cells = load_fixture()
        cls.params = reference_project()
        cls.agg = compute_branches([reference_branch()], cls.params)
        cls.exit_sand = compute_exit_sand(reference_exit_nodes(), cls.params)
        cls.pits = compute_pits(reference_pit(), cls.params)
        cls.summary = compute_earthworks_summary(
            cls.agg, cls.exit_sand,
            hand_trench_count=92,
            hand_trench_section_length=4.0,
            hand_trench_width=0.7,
            params=cls.params,
        )

    def _row(self, dn: int):
        for r in self.agg.rows:
            if r.dn_mm == dn:
                return r
        raise AssertionError(f"row for DN={dn} not found")

    # --- B7 — sand per gnb exit -------------------------------------------

    def test_B7_sand_per_gnb_exit(self):
        assert_close(self, self.params.sand_per_gnb_exit, self.cells["B7"]["value"], "B7")

    # --- G / I (length math) — точно совпадают с xlsm ---------------------

    def test_lengths_per_row(self):
        for dn, addr_g, addr_i in [
            (63,  "G10", "I10"),
            (90,  "G11", "I11"),
            (225, "G12", "I12"),
            (160, "G13", "I13"),
            (110, "G14", "I14"),
        ]:
            r = self._row(dn)
            assert_close(self, r.total_length, self.cells[addr_g]["value"], addr_g)
            assert_close(self, r.open_length,  self.cells[addr_i]["value"], addr_i)

    # --- J (pipe volume) — точно совпадает с xlsm -------------------------

    def test_pipe_volume_per_row(self):
        for dn, addr in [(63, "J10"), (90, "J11"), (225, "J12"), (160, "J13"), (110, "J14")]:
            r = self._row(dn)
            assert_close(self, r.pipe_volume, self.cells[addr]["value"], addr)

    def test_pipe_volume_total(self):
        assert_close(self, self.agg.total_pipe_volume, self.cells["J19"]["value"], "J19")

    # --- I ИТОГО ----------------------------------------------------------

    def test_open_length_total(self):
        assert_close(self, self.agg.total_open_length, self.cells["I19"]["value"], "I19")

    # --- C53 snow clearing (using winter season → all length counts) -----

    def test_snow_clearing(self):
        assert_close(self, self.summary.snow_clearing, self.cells["C53"]["value"], "C53")

    # --- exit sand block ---------------------------------------------------

    def test_exit_sand_total(self):
        assert_close(self, self.exit_sand.total, self.cells["C37"]["value"], "C37")

    # --- pits --------------------------------------------------------------

    def test_pit_block(self):
        p = self.pits
        assert_close(self, p.working, self.cells["C46"]["value"], "C46")
        assert_close(self, p.excavation, self.cells["C48"]["value"], "C48")
        assert_close(self, p.backfill, self.cells["C49"]["value"], "C49")
        assert_close(self, p.shoring_area, self.cells["C50"]["value"], "C50")

    # --- hand trench (C54-C56) — точно совпадает -------------------------

    def test_hand_trench(self):
        assert_close(self, self.summary.hand_trench_excavation, self.cells["C54"]["value"], "C54")
        assert_close(self, self.summary.hand_trench_backfill,    self.cells["C55"]["value"], "C55")
        assert_close(self, self.summary.hand_trench_compact,     self.cells["C56"]["value"], "C56")

    # --- balance is zero ---------------------------------------------------

    def test_balance_zero(self):
        # Баланс должен ОБЯЗАТЕЛЬНО сходиться (выемка == засыпка + вывоз).
        self.assertAlmostEqual(self.summary.balance, 0.0, places=6,
            msg=f"F63 balance = {self.summary.balance}, должен быть 0")


# ============================================================================
# Tests — values that DRIFT from xlsm (documented in module docstring)
# ============================================================================

class TestEarthworkRefactoredFormula(unittest.TestCase):
    """Эти ячейки отличаются от xlsm потому что лечим баги исходного.

    K row 12-14 в xlsm всегда добавлял sand_pad к глубине; теперь это
    добавляется только если sand_cushion_required=True. Для эталона
    (sand_cushion_required=False) это даёт расхождение −5% по K rows 12-14.
    """

    @classmethod
    def setUpClass(cls):
        cls.params = reference_project()
        cls.agg = compute_branches([reference_branch()], cls.params)
        cls.exit_sand = compute_exit_sand(reference_exit_nodes(), cls.params)
        cls.summary = compute_earthworks_summary(
            cls.agg, cls.exit_sand,
            hand_trench_count=92, hand_trench_section_length=4.0, hand_trench_width=0.7,
            params=cls.params,
        )

    def _row(self, dn: int):
        for r in self.agg.rows:
            if r.dn_mm == dn:
                return r
        raise AssertionError(f"row for DN={dn} not found")

    # --- K (soil volume) per row — расчёт вручную -----------------------

    def test_K_row_DN63(self):
        # I=1318.4, depth_total=1.8+0.063=1.863, bw=0.7
        expected = 1318.4 * (1.8 + 0.063) * 0.7
        assert_close(self, self._row(63).soil_volume, expected, "K row DN63")

    def test_K_row_DN90(self):
        expected = 2128.3 * (1.8 + 0.090) * 0.7
        assert_close(self, self._row(90).soil_volume, expected, "K row DN90")

    def test_K_row_DN225(self):
        expected = 5.0 * (1.8 + 0.225) * 0.7
        assert_close(self, self._row(225).soil_volume, expected, "K row DN225")

    def test_K_row_DN160(self):
        expected = 2110.4 * (1.8 + 0.160) * 0.7
        assert_close(self, self._row(160).soil_volume, expected, "K row DN160")

    def test_K_row_DN110(self):
        expected = 662.7 * (1.8 + 0.110) * 0.7
        assert_close(self, self._row(110).soil_volume, expected, "K row DN110")

    def test_K_total(self):
        # Сумма всех 5 строк = ~8323.65 (xlsm имел 8518.12)
        self.assertAlmostEqual(self.agg.total_soil_volume, 8323.6525, places=3)

    def test_L_total_zero(self):
        # sand_side_backfill_required=False → L=0 везде
        self.assertEqual(self.agg.total_sand_side, 0.0)

    def test_M_total_zero(self):
        # sand_cushion_required=False → M=0 везде
        self.assertEqual(self.agg.total_sand_cushion, 0.0)

    def test_C57_excavator(self):
        # (K - C54) × 1.05 = (8323.6525 - 463.68) × 1.05
        expected = (8323.6525 - 463.68) * 1.05
        assert_close(self, self.summary.excavator_excavation, expected, "C57", rel_tol=1e-6)


# ============================================================================
# Tests — non-reference scenarios (sand_cushion, slopes, sections)
# ============================================================================

class TestEarthworkScenarios(unittest.TestCase):
    """Сценарии не из эталона: подушка из песка, откосы, разные участки."""

    def test_sand_cushion_increases_depth(self):
        """С песчаной подушкой K растёт на sand_pad."""
        params = ProjectParams(depth=1.5, snow_thickness=0.5, sand_pad=0.1)
        s_no  = Section(name="no",  soil=SoilType.SAND, season="winter",
                        sand_cushion_required=False)
        s_yes = Section(name="yes", soil=SoilType.SAND, season="winter",
                        sand_cushion_required=True)
        pipes = [PipeRow(dn_mm=160, length_grp1=100.0)]
        agg_no  = compute_branches([Branch(id="b", section=s_no,  pipes=pipes)], params)
        agg_yes = compute_branches([Branch(id="b", section=s_yes, pipes=pipes)], params)

        depth_no  = 1.5 + 0.16              # без подушки
        depth_yes = 1.5 + 0.16 + 0.1        # с подушкой
        self.assertAlmostEqual(agg_no.rows[0].soil_volume,  100.0 * depth_no  * 0.7, places=6)
        self.assertAlmostEqual(agg_yes.rows[0].soil_volume, 100.0 * depth_yes * 0.7, places=6)

    def test_summer_slope_increases_volume(self):
        """Летом песок при глубине 1.5+0.16=1.66 м → 1:1.0 (потому что 1.66 > 1.5)"""
        # depth=1.5 + DN 160 = 1.66 → таблица СНиП попадает в строку h ≤ 3.0 → slope=1.0
        params = ProjectParams(depth=1.5, snow_thickness=0.5, sand_pad=0.1)
        s_winter = Section(name="w", soil=SoilType.SAND, season="winter")
        s_summer = Section(name="s", soil=SoilType.SAND, season="summer")
        pipes = [PipeRow(dn_mm=160, length_grp1=100.0)]
        agg_w = compute_branches([Branch(id="b", section=s_winter, pipes=pipes)], params)
        agg_s = compute_branches([Branch(id="b", section=s_summer, pipes=pipes)], params)

        depth_total = 1.5 + 0.16  # 1.66 м
        # winter: вертикально, площадь = 0.7 × 1.66
        # summer: песок при 1.66 м → попадает в "до 3 м" → slope=1.0; площадь = (0.7+1.66)×1.66
        rect_area = 0.7 * depth_total
        trap_area = (0.7 + depth_total * 1.0) * depth_total
        self.assertAlmostEqual(agg_w.rows[0].soil_volume, 100.0 * rect_area, places=6)
        self.assertAlmostEqual(agg_s.rows[0].soil_volume, 100.0 * trap_area, places=6)
        # Объём при откосе должен быть существенно больше
        self.assertGreater(agg_s.rows[0].soil_volume, agg_w.rows[0].soil_volume * 1.5)

    def test_summer_no_snow_clearing(self):
        """Летом расчистка от снега = 0 (даже если snow_thickness > 0)."""
        params = ProjectParams(depth=1.5, snow_thickness=0.5)
        s = Section(name="s", soil=SoilType.SAND, season="summer")
        agg = compute_branches([Branch(id="b", section=s, pipes=[
            PipeRow(dn_mm=110, length_grp1=500.0),
        ])], params)
        self.assertEqual(agg.snow_clearing_volume, 0.0)
        self.assertEqual(agg.snow_open_length, 0.0)

    def test_mixed_seasons(self):
        """Один проект, две ветки: одна зимняя, одна летняя — снег только за зимнюю."""
        params = ProjectParams(depth=1.5, snow_thickness=0.5)
        winter = Section(name="w", soil=SoilType.SAND, season="winter")
        summer = Section(name="s", soil=SoilType.SAND, season="summer")
        agg = compute_branches([
            Branch(id="b1", section=winter, pipes=[PipeRow(dn_mm=110, length_grp1=500.0)]),
            Branch(id="b2", section=summer, pipes=[PipeRow(dn_mm=110, length_grp1=300.0)]),
        ], params)
        # snow_open_length = только b1's length = 500
        self.assertEqual(agg.snow_open_length, 500.0)
        # snow_volume = 500 × 3.2 × 0.5 = 800
        self.assertEqual(agg.snow_clearing_volume, 800.0)

    def test_forced_overrides_norms(self):
        """forced_slope_ratio / forced_bottom_width перебивают норматив."""
        params = ProjectParams(depth=1.5, snow_thickness=0.5)
        s = Section(name="custom", soil=SoilType.CLAY, season="summer",
                    forced_slope_ratio=0.75, forced_bottom_width=1.0)
        agg = compute_branches([Branch(id="b", section=s, pipes=[
            PipeRow(dn_mm=110, length_grp1=100.0),
        ])], params)
        r = agg.rows[0]
        self.assertEqual(r.profile.slope_ratio, 0.75)
        self.assertEqual(r.profile.bottom_width, 1.0)


if __name__ == "__main__":
    unittest.main()
