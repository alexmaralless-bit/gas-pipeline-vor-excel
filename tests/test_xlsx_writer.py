"""
Tests for pipe_engine.xlsx_writer.

Two kinds of assertions:

1) **Formula presence** — cells that should hold formulas actually
   contain strings starting with "=". This is the core invariant
   ("formulas on output, not pre-computed numbers").

2) **Formula correctness** — read the file back with openpyxl
   data_only=False and compare formula strings to expectations.
   Numeric correctness is already covered by ``test_earthwork.py``
   (the formulas mirror the pure-Python compute path).

For convenience we also generate a sample xlsx in /tmp that the user
can open in Excel manually for visual inspection (test_demo).
"""
from __future__ import annotations

import math
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from pipe_engine.earthwork import ExitSandNode, PitParams, ProjectParams
from pipe_engine.trench_profile import Branch, JointPipePair, PipeRow, Section, SoilType
from pipe_engine.cvor_writer import (
    SHEET_CVOR_CHECKS,
    SHEET_CVOR_FILES,
    SHEET_CVOR_LINKS,
    SHEET_CVOR_META,
    SHEET_CVOR_RESOURCES,
    SHEET_CVOR_SECTIONS,
    SHEET_CVOR_SIGNATURES,
    SHEET_CVOR_WORKS,
)
from pipe_engine.xlsx_writer import (
    IDX_BOTTOM_WIDTH,
    IDX_DEPTH,
    IDX_SAND_CUSHION_FLAG,
    IDX_SAND_PAD,
    IDX_SAND_PER_GNB_EXIT,
    IDX_SEASON_WINTER_FLAG,
    IDX_SLOPE_RATIO,
    IDX_SNOW_THICKNESS,
    SHEET_INPUTS,
    SHEET_TRACE,
    XlsxRenderInputs,
    write_xlsx,
)


# ============================================================================
# Small fixture: single-section reference-like project
# ============================================================================

def reference_inputs() -> XlsxRenderInputs:
    section = Section(name="demo", soil=SoilType.SAND, season="winter")
    branches = [
        Branch(id="main_г1_stage3", section=section, pipes=[
            PipeRow(dn_mm=63,  length_grp1=867.4),
            PipeRow(dn_mm=90,  length_grp1=821.9),
            PipeRow(dn_mm=110, length_grp1=2330.8),
            PipeRow(dn_mm=160, length_grp1=1445.9),
            PipeRow(dn_mm=225, length_grp1=665.9),
            PipeRow(dn_mm=315, length_grp1=400.4),
        ]),
        Branch(id="consumer_inlets_г1_stage3", section=section, pipes=[
            PipeRow(dn_mm=63, length_grp1=2410.4),
        ]),
    ]
    exit_sand = [
        ExitSandNode("ЗУ_кран",        count=21,  sand_units=4),
        ExitSandNode("КТ_по_трассе",   count=385, sand_units=1),
    ]
    return XlsxRenderInputs(
        project_params=ProjectParams(depth=1.8, snow_thickness=0.5, sand_pad=0.1),
        branches=branches,
        exit_sand_nodes=exit_sand,
        pits=PitParams(working_pit_count=72),
        hand_trench_count=92,
        bottom_width=0.7,
        slope_ratio=0.0,
        sand_cushion_required=False,
        season_winter=True,
    )


# ============================================================================
# Helpers
# ============================================================================

def write_and_open(inputs: XlsxRenderInputs):
    """Write to a temp file, load back with formulas (not values)."""
    tmp = Path(tempfile.mkdtemp()) / "out.xlsx"
    write_xlsx(inputs, tmp)
    wb = load_workbook(tmp, data_only=False)
    return wb, tmp


# ============================================================================
# Tests
# ============================================================================

class TestInputsSheet(unittest.TestCase):
    def setUp(self):
        self.wb, _ = write_and_open(reference_inputs())
        self.inp = self.wb[SHEET_INPUTS]

    def test_depth_is_literal(self):
        self.assertEqual(self.inp.cell(IDX_DEPTH, 2).value, 1.8)

    def test_sand_pad_is_literal(self):
        self.assertEqual(self.inp.cell(IDX_SAND_PAD, 2).value, 0.1)

    def test_snow_thickness_is_literal(self):
        self.assertEqual(self.inp.cell(IDX_SNOW_THICKNESS, 2).value, 0.5)

    def test_bottom_width_is_literal(self):
        self.assertEqual(self.inp.cell(IDX_BOTTOM_WIDTH, 2).value, 0.7)

    def test_slope_ratio_is_literal(self):
        self.assertEqual(self.inp.cell(IDX_SLOPE_RATIO, 2).value, 0.0)

    def test_sand_cushion_flag(self):
        self.assertEqual(self.inp.cell(IDX_SAND_CUSHION_FLAG, 2).value, 0)

    def test_season_winter_flag(self):
        self.assertEqual(self.inp.cell(IDX_SEASON_WINTER_FLAG, 2).value, 1)

    def test_sand_per_exit_is_formula(self):
        v = self.inp.cell(IDX_SAND_PER_GNB_EXIT, 2).value
        self.assertIsInstance(v, str)
        self.assertTrue(v.startswith("="))
        # references depth cell
        self.assertIn(f"$B${IDX_DEPTH}", v)


class TestTraceSheetFormulas(unittest.TestCase):
    """All per-row computed cells must be formulas, not numbers."""

    @classmethod
    def setUpClass(cls):
        cls.wb, _ = write_and_open(reference_inputs())
        cls.trace = cls.wb[SHEET_TRACE]

    def test_header_present(self):
        self.assertEqual(self.trace.cell(1, 1).value, "Ветка")
        self.assertEqual(self.trace.cell(1, 2).value, "Ду, мм")

    def test_dn_is_literal(self):
        # Row 2 = first pipe (main DN63)
        self.assertEqual(self.trace.cell(2, 2).value, 63)

    def test_joint_length_is_formula(self):
        # E (col 5) = joint length. No joint pairs in reference_inputs → =0 (still a formula)
        v = self.trace.cell(2, 5).value
        self.assertIsInstance(v, str)
        self.assertTrue(v.startswith("="))

    def test_open_length_is_formula(self):
        # F (col 6) = open length = C - D - E
        v = self.trace.cell(2, 6).value
        self.assertIsInstance(v, str)
        self.assertTrue(v.startswith("="))
        self.assertIn("C2-D2-E2", v)

    def test_pipe_volume_formula_uses_PI(self):
        # G (col 7) = pipe volume J
        v = self.trace.cell(2, 7).value
        self.assertIn("PI()", v)
        self.assertIn("B2/1000", v)
        self.assertIn("/4", v)
        self.assertIn("F2", v)  # references open length (F)

    def test_soil_volume_formula_references_inputs_sheet(self):
        # H (col 8) = soil volume K
        v = self.trace.cell(2, 8).value
        self.assertIn(f"{SHEET_INPUTS}!$B${IDX_DEPTH}", v)
        self.assertIn(f"{SHEET_INPUTS}!$B${IDX_SAND_PAD}", v)
        self.assertIn(f"{SHEET_INPUTS}!$B${IDX_BOTTOM_WIDTH}", v)
        self.assertIn(f"{SHEET_INPUTS}!$B${IDX_SLOPE_RATIO}", v)
        self.assertIn(f"{SHEET_INPUTS}!$B${IDX_SAND_CUSHION_FLAG}", v)
        # References the row's open-length cell (now F)
        self.assertIn("F2", v)

    def test_sand_side_conditional_on_cushion_flag(self):
        # I (col 9) = sand at sides L
        v = self.trace.cell(2, 9).value
        self.assertIn("IF(", v)
        self.assertIn(f"{SHEET_INPUTS}!$B${IDX_SAND_CUSHION_FLAG}", v)

    def test_sand_cushion_conditional_on_cushion_flag(self):
        # J (col 10) = sand cushion M
        v = self.trace.cell(2, 10).value
        self.assertIn("IF(", v)
        self.assertIn(f"{SHEET_INPUTS}!$B${IDX_SAND_CUSHION_FLAG}", v)

    def test_итого_row_has_SUM_formulas(self):
        # We have 7 pipe rows (6 main + 1 inlets), so ИТОГО is row 9
        итого_row = 9
        self.assertEqual(self.trace.cell(итого_row, 1).value, "ИТОГО")
        # Cols 3..10: trace_len, gnb_len, joint_len, open_len, pipe_vol, soil_vol, sand_side, sand_cushion
        for col in (3, 4, 5, 6, 7, 8, 9, 10):
            v = self.trace.cell(итого_row, col).value
            self.assertIsInstance(v, str, f"col {col} should be a formula")
            self.assertTrue(v.startswith("=SUM("), f"col {col}: {v!r}")


class TestEarthworksSummaryFormulas(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.wb, _ = write_and_open(reference_inputs())
        cls.trace = cls.wb[SHEET_TRACE]

    def test_snow_clearing_conditional_on_season(self):
        # Find the snow row by label
        for r in range(1, self.trace.max_row + 1):
            label = self.trace.cell(r, 1).value
            if label and "снег" in str(label).lower():
                v = self.trace.cell(r, 3).value
                self.assertIn("IF(", v)
                self.assertIn(f"{SHEET_INPUTS}!$B${IDX_SEASON_WINTER_FLAG}", v)
                self.assertIn(f"{SHEET_INPUTS}!$B${IDX_SNOW_THICKNESS}", v)
                return
        self.fail("snow-clearing row not found")

    def test_balance_difference_row_present(self):
        # Last meaningful row should be "Разница" with a formula
        for r in range(1, self.trace.max_row + 1):
            label = self.trace.cell(r, 1).value
            if label and "Разница" in str(label):
                v = self.trace.cell(r, 3).value
                self.assertIsInstance(v, str)
                self.assertTrue(v.startswith("="))
                return
        self.fail("Balance Разница row not found")

    def test_excess_soil_is_berm_not_tbo_haul(self):
        labels = [str(self.trace.cell(r, 1).value or "") for r in range(1, self.trace.max_row + 1)]
        self.assertTrue(any("Устройство валика" in label for label in labels))
        self.assertFalse(any("ТБО" in label for label in labels))


class TestExitSandBlock(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.wb, _ = write_and_open(reference_inputs())
        cls.trace = cls.wb[SHEET_TRACE]

    def test_exit_sand_node_has_volume_formula(self):
        for r in range(1, self.trace.max_row + 1):
            if self.trace.cell(r, 1).value == "ЗУ_кран":
                v = self.trace.cell(r, 4).value
                self.assertIn(f"{SHEET_INPUTS}!$B${IDX_SAND_PER_GNB_EXIT}", v)
                self.assertIn(f"B{r}*C{r}", v)
                return
        self.fail("ЗУ_кран row not found")


class TestSchemaInvariants(unittest.TestCase):
    """Cross-cutting checks that must hold."""

    def test_no_numeric_cell_in_computed_columns(self):
        """Per-row computed columns (E..J) must NEVER contain numbers — only
        formula strings. This is the core 'formulas not numbers' contract.

        After the joint-pipe column was added, the computed columns are:
        E=joint_len, F=open_len, G=pipe_vol, H=soil_vol, I=sand_side, J=sand_cushion.
        """
        wb, _ = write_and_open(reference_inputs())
        trace = wb[SHEET_TRACE]
        # Data rows are 2..(2 + 7 - 1) = 2..8
        for r in range(2, 9):
            for col in (5, 6, 7, 8, 9, 10):  # E..J
                v = trace.cell(r, col).value
                self.assertIsInstance(v, str,
                    f"Row {r} col {col}: expected formula string, got {type(v).__name__}: {v!r}")
                self.assertTrue(v.startswith("="),
                    f"Row {r} col {col}: expected '=...' formula, got {v!r}")

    def test_no_orphaned_cell_addresses(self):
        """Catches the C34='D17' bug class: any cell whose VALUE is a string
        looking exactly like a cell address (e.g. 'D17', 'AB123') almost
        certainly forgot the '=' prefix and would be interpreted by Excel
        as literal text, causing #ЗНАЧ! when used arithmetically.

        Skip data-only sheets (Каталог) where literal column-header strings
        like "DN2" can false-positive (Excel column DN exists, so "DN2"
        regex-matches as an address even though it's just a header).
        """
        import re
        CELL_ADDR_RE = re.compile(r"^[A-Z]{1,3}\d+$")
        DATA_ONLY_SHEETS = {
            "Каталог",
            "Тег",
            SHEET_CVOR_META,
            SHEET_CVOR_FILES,
            SHEET_CVOR_SECTIONS,
            SHEET_CVOR_WORKS,
            SHEET_CVOR_RESOURCES,
            SHEET_CVOR_LINKS,
            SHEET_CVOR_SIGNATURES,
            SHEET_CVOR_CHECKS,
        }   # reference/staging data, may contain XML IDs like W00001
        wb, _ = write_and_open(reference_inputs())
        bugs = []
        for sname in wb.sheetnames:
            if sname in DATA_ONLY_SHEETS:
                continue
            ws = wb[sname]
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and CELL_ADDR_RE.match(v):
                        addr = ws.cell(r, c).coordinate
                        bugs.append(f"  [{sname}] {addr} = {v!r} — missing '=' prefix?")
        if bugs:
            self.fail(
                "Found cell(s) containing literal cell-address strings (missing '=' "
                "prefix would make Excel treat them as text, causing #ЗНАЧ!):\n"
                + "\n".join(bugs)
            )

    def test_balance_row_references_all_components(self):
        """The 'Разница' row should subtract excavation - backfill - exit_sand - pipe_volume.
        Each of those four components must be a formula (not literal). Catches both
        the C34 bug (exit_sand as text) and the structurally-similar bugs above it."""
        wb, _ = write_and_open(reference_inputs())
        trace = wb[SHEET_TRACE]
        # Find the four balance-component rows by their label
        labels_required = ["Выемка", "Засыпка", "Песок выходов", "Объём труб"]
        found = {}
        for r in range(1, trace.max_row + 1):
            label = trace.cell(r, 1).value
            if not label:
                continue
            for needle in labels_required:
                if needle in str(label) and needle not in found:
                    found[needle] = (r, trace.cell(r, 3).value)
        self.assertEqual(set(found.keys()), set(labels_required),
                         f"Missing balance rows: {set(labels_required) - set(found.keys())}")
        for needle, (r, v) in found.items():
            self.assertIsInstance(v, str,
                f"Balance row '{needle}' (row {r}) value should be a formula string, "
                f"got {type(v).__name__}: {v!r}")
            self.assertTrue(v.startswith("="),
                f"Balance row '{needle}' (row {r}) should start with '=', got {v!r}")


class TestDemoWorkbook(unittest.TestCase):
    """Generate a real xlsx file the user can open in Excel for visual inspection.
    Always passes if write succeeds — the demo path is logged so the user can find it."""

    def test_write_demo(self):
        out = Path(tempfile.gettempdir()) / "pipe_engine_demo.xlsx"
        write_xlsx(reference_inputs(), out)
        self.assertTrue(out.exists())
        self.assertGreater(out.stat().st_size, 2000, "xlsx should be at least 2KB")
        print(f"\n[DEMO] xlsx written to: {out}")


class TestCvorSheets(unittest.TestCase):
    def test_cvor_xml_staging_sheets_are_emitted(self):
        inputs = reference_inputs()
        inputs.source_dwg_path = r"C:\Projects\Gas\trace.dwg"
        inputs.source_extracted_at_utc = "2026-05-31T12:00:00Z"
        wb, _ = write_and_open(inputs)

        for sheet in (
            SHEET_CVOR_META,
            SHEET_CVOR_FILES,
            SHEET_CVOR_SECTIONS,
            SHEET_CVOR_WORKS,
            SHEET_CVOR_RESOURCES,
            SHEET_CVOR_LINKS,
            SHEET_CVOR_SIGNATURES,
            SHEET_CVOR_CHECKS,
        ):
            self.assertIn(sheet, wb.sheetnames)

        files = wb[SHEET_CVOR_FILES]
        self.assertEqual(files.cell(2, 1).value, "F001")
        self.assertEqual(files.cell(2, 2).value, "DWG")
        self.assertEqual(files.cell(2, 3).value, "trace.dwg")
        self.assertEqual(files.cell(2, 5).value, "Model")

    def test_cvor_works_reference_vr_and_catalog_codes(self):
        inputs = reference_inputs()
        inputs.source_dwg_path = r"C:\Projects\Gas\trace.dwg"
        wb, _ = write_and_open(inputs)
        works = wb[SHEET_CVOR_WORKS]

        self.assertGreater(works.max_row, 2)
        self.assertEqual(works.cell(1, 1).value, "WorkID")
        self.assertEqual(works.cell(2, 5).value[:6], "='ВР'!")
        self.assertEqual(works.cell(2, 7).value[:6], "='ВР'!")

        types = [works.cell(r, 4).value for r in range(2, works.max_row + 1)]
        self.assertIn("материал", types)
        self.assertIn("работа", types)

    def test_cvor_links_cover_every_work(self):
        inputs = reference_inputs()
        inputs.source_dwg_path = r"C:\Projects\Gas\trace.dwg"
        wb, _ = write_and_open(inputs)
        works = wb[SHEET_CVOR_WORKS]
        links = wb[SHEET_CVOR_LINKS]

        work_ids = [works.cell(r, 1).value for r in range(2, works.max_row + 1)]
        link_work_ids = [links.cell(r, 2).value for r in range(2, links.max_row + 1)]
        self.assertEqual(work_ids, link_work_ids)
        self.assertTrue(all(links.cell(r, 3).value == "F001" for r in range(2, links.max_row + 1)))


class TestJointPipeBlock(unittest.TestCase):
    """Joint-pipe block is rendered only when joint_pipe_pairs is non-empty,
    and trace-row column E uses SUMIF to pull joint length per DN."""

    def _inputs_with_joint(self) -> XlsxRenderInputs:
        section = Section(name="t", soil=SoilType.SAND, season="winter")
        branches = [Branch(id="m", section=section, pipes=[
            PipeRow(dn_mm=110, length_grp1=1000.0),
            PipeRow(dn_mm=63,  length_grp1=500.0),
        ])]
        # Joint pair: 100 m of (DN 110 + DN 63) laid jointly
        joint_pairs = [JointPipePair(id="j1", dn1_mm=110, dn2_mm=63, length_m=100.0)]
        return XlsxRenderInputs(
            project_params=ProjectParams(depth=1.8, snow_thickness=0.5, sand_pad=0.1),
            branches=branches, exit_sand_nodes=[], pits=None,
            hand_trench_count=0, joint_pipe_pairs=joint_pairs,
        )

    def test_joint_block_emitted_when_pairs_present(self):
        wb, _ = write_and_open(self._inputs_with_joint())
        ws = wb[SHEET_TRACE]
        # Find the joint block section title
        found = False
        for r in range(1, ws.max_row + 1):
            if str(ws.cell(r, 1).value or "").startswith("СОВМЕСТНАЯ"):
                found = True
                break
        self.assertTrue(found, "Joint block section title not found")

    def test_joint_block_absent_when_pairs_empty(self):
        wb, _ = write_and_open(reference_inputs())  # joint_pipe_pairs=[]
        ws = wb[SHEET_TRACE]
        for r in range(1, ws.max_row + 1):
            label = str(ws.cell(r, 1).value or "")
            self.assertFalse(label.startswith("СОВМЕСТНАЯ"),
                f"Joint block must NOT appear when no joint pairs; found at row {r}")

    def test_trace_joint_col_uses_SUMIF_when_pairs_present(self):
        wb, _ = write_and_open(self._inputs_with_joint())
        ws = wb[SHEET_TRACE]
        v = ws.cell(2, 5).value  # E2 = joint length for DN 110
        self.assertIn("SUMIF", v)
        # Should reference both DN1 column (B) and DN2 column (C) of joint block
        self.assertIn("B2", v)  # this row's DN

    def test_trace_joint_col_is_zero_formula_when_no_pairs(self):
        wb, _ = write_and_open(reference_inputs())
        ws = wb[SHEET_TRACE]
        v = ws.cell(2, 5).value
        self.assertEqual(v, "=0", f"E2 should be '=0' when no joint pairs, got {v!r}")


class TestFormulaEvalMatchesPython(unittest.TestCase):
    """Closes the loop: Excel evaluating our xlsx formulas must produce the
    SAME numbers as the pure-Python compute path (pipe_engine.earthwork).

    Uses the `formulas` package (Eurostat-grade Excel evaluator) to actually
    compute every formula in the generated xlsx and compare to
    compute_branches/compute_earthworks_summary outputs.

    Combined with test_earthwork.py (which validates Python vs xlsm cached
    values), this transitively proves:
        xlsm cached ≡ Python ≡ Excel evaluating my formulas
    to within the documented Step 2.5 drift.

    Skips with a clear message if `formulas` package is not installed
    (it's a dev-only verification dep, not a runtime dep).
    """

    @classmethod
    def setUpClass(cls):
        try:
            import formulas  # noqa: F401
            cls.formulas = formulas
        except ImportError:
            raise unittest.SkipTest(
                "formulas package not installed — run `pip install formulas` "
                "to enable end-to-end xlsx formula verification."
            )

    def _build_reference_snapshot(self):
        """Same inputs as test_earthwork.py — эталонной книги row 10-14 + B30-B36."""
        from pipe_engine.earthwork import (
            compute_branches, compute_earthworks_summary, compute_exit_sand,
        )
        section = Section(name="demo", soil=SoilType.SAND, season="winter",
                          sand_cushion_required=False,
                          sand_side_backfill_required=False)
        branches = [Branch(id="main", section=section, pipes=[
            PipeRow(dn_mm=63,  length_grp1=968.9,  length_other=587.0,  gnb_length=237.5),
            PipeRow(dn_mm=90,  length_grp1=1277.9, length_other=1004.0, gnb_length=153.6),
            PipeRow(dn_mm=225, length_grp1=5.0,    length_other=0.0,    gnb_length=0.0),
            PipeRow(dn_mm=160, length_grp1=836.6,  length_other=1877.0, gnb_length=603.2),
            PipeRow(dn_mm=110, length_grp1=684.9,  length_other=241.0,  gnb_length=263.2),
        ])]
        exit_sand = [
            ExitSandNode("ЗУ надземно",                21,  4),
            ExitSandNode("Узел А и Б",                155,  2),
            ExitSandNode("КТ",                        156,  1),
            ExitSandNode("Пересеч инж комм в футляре", 92,  1),
            ExitSandNode("Прокладка в футляре ГНБ",    44,  1),
        ]
        params = ProjectParams(depth=1.8, snow_thickness=0.5, sand_pad=0.1)
        inputs = XlsxRenderInputs(
            project_params=params, branches=branches, exit_sand_nodes=exit_sand,
            pits=PitParams(working_pit_count=72), hand_trench_count=92,
            bottom_width=0.7, slope_ratio=0.0,
            sand_cushion_required=False, season_winter=True,
        )
        agg = compute_branches(branches, params)
        es  = compute_exit_sand(exit_sand, params)
        summary = compute_earthworks_summary(agg, es, hand_trench_count=92,
            hand_trench_section_length=4.0, hand_trench_width=0.7, params=params)
        return inputs, agg, es, summary

    def test_all_key_cells_match_python_compute(self):
        inputs, agg, es, summary = self._build_reference_snapshot()
        out = Path(tempfile.mkdtemp()) / "reference_snapshot_eval.xlsx"
        write_xlsx(inputs, out)

        xl = self.formulas.ExcelModel().loads(str(out)).finish()
        calc = xl.calculate()
        fname = out.name

        def fetch(sheet: str, cell: str):
            # formulas uppercases sheet names internally
            key = f"'[{fname}]{sheet.upper()}'!{cell}"
            if key not in calc:
                return None
            v = calc[key].value
            try: return float(v[0][0])
            except Exception:
                try: return float(v)
                except Exception: return None

        # Each tuple: (label, sheet, cell_in_xlsx, expected_python_value)
        # Cell addresses depend on the layout produced by build_trace_sheet
        # + append_exit_sand_block (inline) + append_earthworks_summary.
        # For 5 pipe rows: data 2..6, ИТОГО r7, exit-sand 9..16, summary 18..28, balance 30..35.
        # After adding joint-pipe column E, COLs shifted: G→H (soil), F→G (pipe vol)
        checks = [
            ("B12 sand_per_exit", "ИсхДанные",   "B12", inputs.project_params.sand_per_gnb_exit),
            ("K DN63",            "Расчет земли", "H2",  agg.rows[0].soil_volume),
            ("K DN90",            "Расчет земли", "H3",  agg.rows[1].soil_volume),
            ("K DN225",           "Расчет земли", "H4",  agg.rows[2].soil_volume),
            ("K DN160",           "Расчет земли", "H5",  agg.rows[3].soil_volume),
            ("K DN110",           "Расчет земли", "H6",  agg.rows[4].soil_volume),
            ("J ИТОГО",           "Расчет земли", "G7",  agg.total_pipe_volume),
            ("K ИТОГО",           "Расчет земли", "H7",  agg.total_soil_volume),
            ("exit_sand ИТОГО",   "Расчет земли", "D16", es.total),
            ("snow",              "Расчет земли", "C19", summary.snow_clearing),
            ("hand trench",       "Расчет земли", "C20", summary.hand_trench_excavation),
            ("excavator",         "Расчет земли", "C23", summary.excavator_excavation),
            ("vyvoz landfill",    "Расчет земли", "C26", summary.landfill_volume),
            ("dozer backfill",    "Расчет земли", "C27", summary.dozer_backfill),
            ("BALANCE",           "Расчет земли", "C35", summary.balance),
        ]
        bugs = []
        for label, sheet, addr, expected in checks:
            actual = fetch(sheet, addr)
            if actual is None:
                bugs.append(f"  {label} [{addr}]: cell not in eval result")
                continue
            if abs(expected) > 1e-9:
                rel = abs(actual - expected) / abs(expected)
                if rel > 1e-6:
                    bugs.append(f"  {label} [{addr}]: python={expected!r} xlsx_eval={actual!r} "
                                f"(rel diff {rel:.2e})")
            else:
                if abs(actual) > 1e-6:
                    bugs.append(f"  {label} [{addr}]: expected ≈0, xlsx_eval={actual!r}")
        if bugs:
            self.fail(
                "xlsx formula evaluation diverges from Python pipeline:\n" + "\n".join(bugs)
            )


class TestJointFormulaEval(unittest.TestCase):
    """Joint-pipe SUMIF formula must correctly find ALL pairs whose DN1 or DN2
    matches each trace row's DN. Off-by-one in the joint-block address range
    would silently miss pairs and produce wrong per-DN joint length.
    """

    @classmethod
    def setUpClass(cls):
        try:
            import formulas  # noqa: F401
            cls.formulas = formulas
        except ImportError:
            raise unittest.SkipTest("formulas package required")

    def test_sumif_finds_all_matching_pairs(self):
        # Two pairs: J1 = (DN 110, DN 63, 250m), J2 = (DN 160, DN 63, 100m).
        # Expected per-DN joint lengths:
        #   DN 110 → 250  (J1 dn1)
        #   DN 63  → 250 + 100 = 350  (J1 dn2 + J2 dn2)
        #   DN 160 → 100  (J2 dn1)
        section = Section(name="t", soil=SoilType.SAND, season="winter")
        branches = [Branch(id="m", section=section, pipes=[
            PipeRow(dn_mm=110, length_grp1=1000.0),
            PipeRow(dn_mm=63,  length_grp1=500.0),
            PipeRow(dn_mm=160, length_grp1=300.0),
        ])]
        joint_pairs = [
            JointPipePair(id="J1", dn1_mm=110, dn2_mm=63, length_m=250.0),
            JointPipePair(id="J2", dn1_mm=160, dn2_mm=63, length_m=100.0),
        ]
        inputs = XlsxRenderInputs(
            project_params=ProjectParams(depth=1.8, snow_thickness=0.5, sand_pad=0.1),
            branches=branches, exit_sand_nodes=[], pits=None, hand_trench_count=0,
            joint_pipe_pairs=joint_pairs,
        )
        out = Path(tempfile.mkdtemp()) / "joint_eval.xlsx"
        write_xlsx(inputs, out)

        xl = self.formulas.ExcelModel().loads(str(out)).finish()
        calc = xl.calculate()
        fname = out.name
        def f(cell):
            key = f"'[{fname}]РАСЧЕТ ЗЕМЛИ'!{cell}"
            if key not in calc: return None
            v = calc[key].value
            try: return float(v[0][0])
            except:
                try: return float(v)
                except: return None

        # Trace data rows: 2, 3, 4 → DN 110, 63, 160
        e2 = f("E2"); e3 = f("E3"); e4 = f("E4")
        self.assertEqual(e2, 250.0, f"DN 110 joint length: expected 250, got {e2!r}")
        self.assertEqual(e3, 350.0, f"DN 63 joint length: expected 350 (J1+J2 dn2), got {e3!r}")
        self.assertEqual(e4, 100.0, f"DN 160 joint length: expected 100 (J2 dn1), got {e4!r}")

        # Trace ИТОГО of joint col
        e5 = f("E5")
        self.assertEqual(e5, 700.0, f"trace E ИТОГО: expected 700, got {e5!r}")

        # Joint block ИТОГО of length col
        e_joint_итого = f("E11")
        self.assertEqual(e_joint_итого, 350.0, f"joint E ИТОГО: expected 350, got {e_joint_итого!r}")

        # BALANCE must be ≈ 0 with joint pipe contributing.
        # Without exit_sand block in this test scenario, balance row is C30
        # (3 trace + ИТОГО + blank + 4 joint + blank + 11 summary + blank + 5 balance).
        # Locate it by label instead of hardcoded row for robustness.
        balance_row = None
        from openpyxl import load_workbook
        wb = load_workbook(str(out), data_only=False)
        ws = wb["Расчет земли"]
        for r in range(1, ws.max_row + 1):
            if "Разница" in str(ws.cell(r, 1).value or ""):
                balance_row = r
                break
        self.assertIsNotNone(balance_row, "Разница row not found")
        balance = f(f"C{balance_row}")
        self.assertAlmostEqual(balance, 0.0, places=6,
            msg=f"Balance must be 0 even with joint laying, got {balance!r}")


if __name__ == "__main__":
    unittest.main()
