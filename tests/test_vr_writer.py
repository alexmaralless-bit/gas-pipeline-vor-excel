"""Tests for pipe_engine.vr_writer — ВР sheet generation."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from pipe_engine.earthwork import ExitSandNode, PitParams, ProjectParams
from pipe_engine.io_schema import BlockRecord
from pipe_engine.node_library import NodeLibrary
from pipe_engine.trench_profile import Branch, PipeRow, Section, SoilType
from pipe_engine.xlsx_writer import SHEET_INPUTS, XlsxRenderInputs, write_xlsx
from pipe_engine.vr_writer import SHEET_VR
from pipe_engine.spec_writer import SHEET_SPEC


def _make_inputs(blocks: list[BlockRecord] | None = None) -> XlsxRenderInputs:
    section = Section(name="t", soil=SoilType.SAND, season="winter")
    branches = [Branch(id="main", section=section, pipes=[
        PipeRow(dn_mm=63, length_grp1=1000.0),
        PipeRow(dn_mm=110, length_grp1=2000.0, gnb_length=200.0),
    ])]
    exit_sand = [
        ExitSandNode("ЗУ", 5, 4),
        ExitSandNode("КТ", 50, 1),
    ]
    return XlsxRenderInputs(
        project_params=ProjectParams(depth=1.8, snow_thickness=0.5, sand_pad=0.1),
        branches=branches, exit_sand_nodes=exit_sand,
        pits=PitParams(working_pit_count=10),
        hand_trench_count=5, sand_cushion_required=False, season_winter=True,
        blocks=blocks or [],
    )


def _write_and_load(inputs: XlsxRenderInputs):
    tmp = Path(tempfile.mkdtemp()) / "out.xlsx"
    write_xlsx(inputs, tmp)
    return load_workbook(tmp, data_only=False), tmp


class TestVrSheetStructure(unittest.TestCase):
    def test_vr_sheet_present(self):
        wb, _ = _write_and_load(_make_inputs())
        self.assertIn(SHEET_VR, wb.sheetnames)

    def test_vr_has_title(self):
        wb, _ = _write_and_load(_make_inputs())
        ws = wb[SHEET_VR]
        self.assertEqual(ws.cell(1, 4).value, "Ведомость объёмов работ")

    def test_vr_section_zemlanye_present(self):
        wb, _ = _write_and_load(_make_inputs())
        ws = wb[SHEET_VR]
        found = False
        for r in range(1, ws.max_row + 1):
            if "Земляные работы" in str(ws.cell(r, 4).value or ""):
                found = True
                break
        self.assertTrue(found, "Section 'Земляные работы' missing from ВР")

    def test_vr_zemlanye_qty_is_formula(self):
        wb, _ = _write_and_load(_make_inputs())
        ws = wb[SHEET_VR]
        # Items start right after section header. First item is "Расчистка трассы от снега"
        for r in range(5, ws.max_row + 1):
            if "снега" in str(ws.cell(r, 4).value or "").lower():
                qty = ws.cell(r, 6).value
                self.assertIsInstance(qty, str, f"F{r} should be formula")
                self.assertTrue(qty.startswith("="), f"F{r}={qty!r}")
                self.assertIn("Расчет земли", qty)
                return
        self.fail("snow row not found")

    def test_vr_uses_excess_soil_berm_instead_of_tbo_haul(self):
        wb, _ = _write_and_load(_make_inputs())
        ws = wb[SHEET_VR]
        names = [str(ws.cell(r, 4).value or "") for r in range(1, ws.max_row + 1)]
        self.assertTrue(any("Устройство валика" in name for name in names))
        self.assertFalse(any("ТБО" in name for name in names))

    def test_vr_work_rows_use_work_catalog_keys(self):
        wb, _ = _write_and_load(_make_inputs())
        ws = wb[SHEET_VR]
        notes = [str(ws.cell(r, 9).value or "") for r in range(1, ws.max_row + 1)]
        self.assertTrue(any("work_key=earth.excavator_trench" in note for note in notes))
        self.assertTrue(any("work_key=pipe.pe_open_laying" in note for note in notes))
        self.assertTrue(any("estimate_norm_code=pending" in note for note in notes))

    def test_vr_pipe_section_uses_ROUNDUP_with_overrun(self):
        wb, _ = _write_and_load(_make_inputs())
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            d = ws.cell(r, 4).value
            if d and "Труба из полиэтилена для газопроводов" in str(d):
                qty = ws.cell(r, 6).value
                self.assertIn("ROUNDUP", qty)
                self.assertIn("ИсхДанные!$B$13", qty)  # overrun factor cell
                self.assertIn("DB lookup: tag='Труба ПЭ'", str(ws.cell(r, 9).value or ""))
                return
        self.fail("no catalog-backed PE pipe row found")

    def test_vr_gnb_section_when_gnb_present(self):
        # Inputs above have DN110 with gnb_length=200
        wb, _ = _write_and_load(_make_inputs())
        ws = wb[SHEET_VR]
        gnb_label_found = False
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            if "ГНБ" in d:
                gnb_label_found = True
                break
        self.assertTrue(gnb_label_found, "ГНБ section should appear when gnb_length>0")

    def test_vr_gnb_section_skipped_when_no_gnb(self):
        section = Section(name="t", soil=SoilType.SAND, season="winter")
        # No gnb anywhere
        branches = [Branch(id="m", section=section, pipes=[
            PipeRow(dn_mm=63, length_grp1=100.0),
        ])]
        inputs = XlsxRenderInputs(
            project_params=ProjectParams(depth=1.8, snow_thickness=0.5, sand_pad=0.1),
            branches=branches, exit_sand_nodes=[], pits=None, hand_trench_count=0,
            sand_cushion_required=False, season_winter=True, blocks=[],
        )
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            self.assertNotIn("Укладка газопроводов подземно закрытым", d,
                msg=f"ГНБ section should be skipped when no GNB; found at r{r}")


class TestVrNodeLibrarySection(unittest.TestCase):
    """Section 5: auto-emit from NodeLibrary for matching blocks."""

    def _block(self, name: str, handle: str = "ABC") -> BlockRecord:
        return BlockRecord(
            handle=handle, block_name_raw=name, block_name_effective=name,
            layer="!ГСН_Арм", position=(0, 0, 0), rotation_rad=0,
            scale=(1, 1, 1), attributes={},
        )

    def test_no_section_when_no_matching_blocks(self):
        # Only blocks that don't match any template
        blocks = [self._block("СО_НЕИЗВЕСТНЫЙ", "1")]
        wb, _ = _write_and_load(_make_inputs(blocks=blocks))
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            self.assertNotIn("Устройство узлов", d, "Nodes section should be absent")

    def test_kt_blocks_emitted(self):
        blocks = [self._block("СО_КТ_по_трассе", str(i)) for i in range(5)]
        wb, _ = _write_and_load(_make_inputs(blocks=blocks))
        ws = wb[SHEET_VR]
        # Group header теперь без F (чтобы не было двойного счёта с компонентами).
        # block_count показан в имени узла как «(всего 5 шт.)».
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            if "контрольной трубки" in d.lower() and "в т. ч." in d:
                self.assertIn("всего 5 шт.", d, f"Expected 'всего 5 шт.' in header D{r}={d!r}")
                self.assertIsNone(ws.cell(r, 6).value,
                    f"Header F{r} should be empty (not sumable), got {ws.cell(r, 6).value!r}")
                return
        self.fail("КТ group header not found in ВР")

    def test_zu_kran_emits_inflated_qty(self):
        blocks = [self._block("СО_ЗУ_кран", str(i)) for i in range(10)]
        wb, _ = _write_and_load(_make_inputs(blocks=blocks))
        ws = wb[SHEET_VR]
        # ZU template has "Соединение неразъемное" with qty=2 per kran → 20 for 10 кранов.
        # Formula now uses literal block_count: F = "=10*2".
        for r in range(1, ws.max_row + 1):
            if "Соединение неразъемное" in str(ws.cell(r, 4).value or ""):
                qty = ws.cell(r, 6).value
                self.assertIsInstance(qty, str, f"F{r} should be formula, got {qty!r}")
                self.assertEqual(qty, "=10*2",
                    f"Expected '=10*2' (block_count×qty_per_unit), got {qty!r}")
                return
        self.fail("Соединение неразъемное row not found")

    def test_real_project_block_aliases_emit_nodes(self):
        blocks = [
            BlockRecord(
                handle="p1",
                block_name_raw="СО_Переход_ПЭ-Сталь",
                block_name_effective="СО_Переход_ПЭ-Сталь",
                layer="!ГСН_Арм",
                position=(0, 0, 0),
                rotation_rad=0,
                scale=(1, 1, 1),
                attributes={"DN": "Переход ПЭ/Сталь DN110x108"},
            ),
            BlockRecord(
                handle="o1",
                block_name_raw="СО_Отвод_90_ПЭ_литой",
                block_name_effective="СО_Отвод_90_ПЭ_литой",
                layer="!ГСН_Арм",
                position=(0, 0, 0),
                rotation_rad=0,
                scale=(1, 1, 1),
                attributes={"DN": "Отвод 90 ПЭ Литой DN110"},
            ),
        ]
        wb, _ = _write_and_load(_make_inputs(blocks=blocks))
        ws = wb[SHEET_VR]
        names = [str(ws.cell(r, 4).value or "") for r in range(1, ws.max_row + 1)]
        self.assertTrue(any("Переход редукционный" in name and "DN110" in name for name in names))
        self.assertTrue(any("Отвод 90" in name and "DN110" in name for name in names))

    def test_futlyar_uses_length_attrs_for_pipe_and_onk_qty(self):
        block = BlockRecord(
            handle="f1",
            block_name_raw="СО_Футляр_ПЭ",
            block_name_effective="СО_Футляр_ПЭ",
            layer="!ГСН_Арм",
            position=(0, 0, 0),
            rotation_rad=0,
            scale=(1, 1, 1),
            attributes={"DN": "Футляр DN250x22.7", "LENGTH": "9.36"},
        )
        wb, _ = _write_and_load(_make_inputs(blocks=[block]))
        ws = wb[SHEET_VR]
        found_pipe = found_onk = found_work = found_onk_work = False
        for r in range(1, ws.max_row + 1):
            name = str(ws.cell(r, 4).value or "")
            qty = ws.cell(r, 6).value
            audit = str(ws.cell(r, 8).value or "")
            note = str(ws.cell(r, 9).value or "")
            if "LENGTH attrs" in audit and "Труба ПЭ" in note:
                self.assertEqual(qty, "=9.36")
                found_pipe = True
            if "ONK per футляр" in audit and "ОНК" in name:
                self.assertEqual(qty, "=ROUND(9.36/2.5,0)+1")
                found_onk = True
            if "Прокладка футляра" in name:
                self.assertEqual(qty, "=9.36")
                found_work = True
            if "Установка опорно-направляющих" in name:
                self.assertEqual(qty, "=ROUND(9.36/2.5,0)+1")
                found_onk_work = True
        self.assertTrue(found_pipe)
        self.assertTrue(found_onk)
        self.assertTrue(found_work)
        self.assertTrue(found_onk_work)


class TestPolylineTrigger(unittest.TestCase):
    """polyline-triggered templates (Узел А by _Вводы layer regex)."""

    def _polyline(self, layer: str, handle: str = "X") -> "PolylineRecord":
        from pipe_engine.io_schema import PolylineRecord
        return PolylineRecord(
            handle=handle, layer=layer, length=10.0,
            start=(0, 0, 0), end=(0, 0, 0), is_closed=False, vertex_count=2,
        )

    def test_uzel_a_emitted_from_inlet_polylines(self):
        inputs = _make_inputs()
        # Add 5 polylines on _Вводы@63 layer → should trigger Узел А template ×5 DN63
        inputs.polylines = [self._polyline("!!ГСН3_Г1_ГП_Вводы@63", str(i)) for i in range(5)]
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        # Header теперь информирует count в имени, F пустой.
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            if "Узел А" in d and "DN63" in d:
                self.assertIn("всего 5 шт.", d, f"Expected '5 шт.' in header D{r}")
                self.assertIsNone(ws.cell(r, 6).value)
                return
        self.fail("Узел А DN63 group not found in ВР")

    def test_uzel_a_emitted_from_project_pe_inlet_polylines(self):
        inputs = _make_inputs()
        inputs.polylines = [self._polyline("!ГСН_Г1_ПЭ_Вводы_@63x5.8", str(i)) for i in range(2)]
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            if "Узел А" in d and "DN63" in d:
                self.assertIn("всего 2 шт.", d)
                return
        self.fail("Узел А DN63 group not found for PE inlet layer")

    def test_uzel_a_skipped_when_no_inlet_polylines(self):
        inputs = _make_inputs()
        inputs.polylines = []
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            self.assertNotIn("Узел А", d)


class TestManualTrigger(unittest.TestCase):
    """manual-triggered templates (ГРПШ, Заземлитель, Врезка, Свеча)."""

    def test_grpsh_emitted_when_count_provided(self):
        inputs = _make_inputs()
        inputs.manual_counts = {"grpsh": 2, "zazemlitel": 0, "vrezka": 0, "svecha": 0}
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            if "ГРПШ" in d and "в т. ч." in d:
                # Header теперь без F-числа; count выведен в имя «(всего 2 шт.)»
                self.assertIn("всего 2 шт.", d, f"Expected '2 шт.' in header D{r}")
                self.assertIsNone(ws.cell(r, 6).value)
                return
        self.fail("ГРПШ header not found")

    def test_zazemlitel_components_multiplied(self):
        inputs = _make_inputs()
        inputs.manual_counts = {"zazemlitel": 4}
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            if "Полоса 40" in str(ws.cell(r, 4).value or ""):
                qty = ws.cell(r, 6).value
                # 5 м per заземлитель × 4 = 20 м. F = "=4*5" с literal block_count.
                self.assertEqual(qty, "=4*5",
                    f"Expected '=4*5' (block_count×qty_per_unit), got {qty!r}")
                return
        self.fail("Полоса 40 component not found")

    def test_zero_count_omits_template(self):
        inputs = _make_inputs()
        inputs.manual_counts = {"grpsh": 0, "zazemlitel": 0, "vrezka": 0}
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            self.assertNotIn("ГРПШ", d)
            self.assertNotIn("Заземлитель", d)
            self.assertNotIn("врезки", d.lower())


class TestNodeComponentDnHint(unittest.TestCase):
    """dn_hint перебивает host_dn узла для catalog DB lookup (Q48)."""

    def test_dn_hint_loaded_from_json(self):
        lib = NodeLibrary.load_default()
        ua = lib.find_by_id("uzel_a_potrebitel")
        self.assertIsNotNone(ua)
        # Узел А: труба Ду50 должна иметь dn_hint=50, труба Ду20 → 20
        труба_ду50 = next(c for c in ua.components if "Ду50" in c.tag)
        труба_ду20 = next(c for c in ua.components if "Ду20" in c.tag)
        self.assertEqual(труба_ду50.dn_hint, 50)
        self.assertEqual(труба_ду20.dn_hint, 20)
        # Кран — должен ОСТАТЬСЯ без dn_hint (использует host_dn)
        кран = next(c for c in ua.components if c.tag.startswith("Кран шаровой изолирующий"))
        self.assertIsNone(кран.dn_hint)

    def test_dn_hint_used_in_db_lookup_when_present(self):
        """DB lookup должен использовать dn_hint компонента, не host_dn узла."""
        from pipe_engine.io_schema import PolylineRecord
        from pipe_engine.catalog_loader import load_catalog

        inputs = _make_inputs()
        # 1 polyline → Узел А с host_dn=63
        inputs.polylines = [PolylineRecord(
            handle="p1", layer="!!ГСН3_Г1_ГП_Вводы@63", length=10.0,
            start=(0, 0, 0), end=(0, 0, 0), is_closed=False, vertex_count=2,
        )]
        # Catalog required for DB-backed name resolution.
        try:
            inputs.catalog = load_catalog()
        except FileNotFoundError:
            self.skipTest("pipe_catalog.db not present in CI")
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]

        # «Труба стальная Ду50» row — lookup должен искать DN=50, не 63.
        # Проверяем через debug-метку в COL_NOTE (I): "DN=50".
        for r in range(1, ws.max_row + 1):
            note = str(ws.cell(r, 9).value or "")
            if "DB lookup" in note and "DN=50" in note:
                # Found dn_hint=50 override
                self.assertIn("dn_hint override", note,
                    f"I{r} should mention dn_hint override: {note!r}")
                return
        self.fail("No component used dn_hint=50 override for Узел А DN63")


class TestGeometricHostDn(unittest.TestCase):
    """host_dn_source='geometry' picks DN from nearest trace polyline (Q5)."""

    def _kt_block(self, x: float, y: float, handle: str = "kt") -> BlockRecord:
        return BlockRecord(
            handle=handle,
            block_name_raw="СО_КТ_по_трассе",
            block_name_effective="СО_КТ_по_трассе",
            layer="!ГСН_Арм",
            position=(x, y, 0),
            rotation_rad=0,
            scale=(1, 1, 1),
            attributes={"DУ": "300"},   # ковра диаметр, не газа — должен быть проигнорирован
        )

    def _trace_polyline(self, dn: int, start_xy, end_xy, handle: str = "tr"):
        from pipe_engine.io_schema import PolylineRecord
        return PolylineRecord(
            handle=handle,
            layer=f"!!ГСН3_Г1_ГП@{dn}",
            length=10.0,
            start=(start_xy[0], start_xy[1], 0),
            end=(end_xy[0], end_xy[1], 0),
            is_closed=False, vertex_count=2,
        )

    def test_kt_picks_dn_from_nearest_trace(self):
        # КТ at (5, 0.5); two traces: DN63 trace AB=(0,0)-(10,0) — distance 0.5
        # DN160 trace far away — distance ~50. КТ should pick DN63.
        blocks = [self._kt_block(5.0, 0.5)]
        inputs = _make_inputs(blocks=blocks)
        inputs.polylines = [
            self._trace_polyline(63, (0, 0), (10, 0), "near"),
            self._trace_polyline(160, (0, 50), (10, 50), "far"),
        ]
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        # Group header for КТ must say "DN63"
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            if "контрольной трубки" in d.lower() and "в т. ч." in d:
                self.assertIn("DN63", d,
                    f"Expected КТ resolved to DN63 from geometry, got header {d!r}")
                return
        self.fail("КТ group header not found")

    def test_kt_far_from_any_trace_stays_none(self):
        # КТ at (5, 100) — both traces > 10м away → host_dn=None
        blocks = [self._kt_block(5.0, 100.0)]
        inputs = _make_inputs(blocks=blocks)
        inputs.polylines = [
            self._trace_polyline(63, (0, 0), (10, 0), "p1"),
        ]
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            if "контрольной трубки" in d.lower() and "в т. ч." in d:
                self.assertNotIn("DN", d,
                    f"Far КТ should not pick any DN, got {d!r}")
                return
        self.fail("КТ group header not found")

    def test_zu_kran_attribute_fallback_to_geometry(self):
        # ZU-кран атрибут «Кран КШ-150с» — не парсится → fallback на geometry → DN110
        block = BlockRecord(
            handle="zu1",
            block_name_raw="СО_ЗУ_кран",
            block_name_effective="СО_ЗУ_кран",
            layer="!ГСН_Арм",
            position=(2.0, 0.5, 0),
            rotation_rad=0,
            scale=(1, 1, 1),
            attributes={"DN": "Кран КШ-150с"},
        )
        inputs = _make_inputs(blocks=[block])
        inputs.polylines = [self._trace_polyline(110, (0, 0), (10, 0))]
        wb, _ = _write_and_load(inputs)
        ws = wb[SHEET_VR]
        for r in range(1, ws.max_row + 1):
            d = str(ws.cell(r, 4).value or "")
            if "Крана надземно" in d and "в т. ч." in d:
                self.assertIn("DN110", d, f"ZU должен взять DN110 fallback, got {d!r}")
                return
        self.fail("ZU group header not found")


class TestNodeLibraryArchitectureV2(unittest.TestCase):
    """Architecture supports 3 trigger types."""

    def test_default_library_has_polyline_template(self):
        from pipe_engine.node_library import NodeLibrary
        lib = NodeLibrary.load_default()
        polyline_templates = lib.templates_by_trigger("polyline_layer")
        self.assertGreaterEqual(len(polyline_templates), 1)
        ua = next((t for t in polyline_templates if t.id == "uzel_a_potrebitel"), None)
        self.assertIsNotNone(ua)
        self.assertIn("Вводы", ua.polyline_layer_pattern)

    def test_default_library_has_manual_templates(self):
        from pipe_engine.node_library import NodeLibrary
        lib = NodeLibrary.load_default()
        manual_templates = lib.templates_by_trigger("manual")
        self.assertGreaterEqual(len(manual_templates), 3)
        manual_keys = {t.manual_count_key for t in manual_templates}
        self.assertIn("grpsh", manual_keys)
        self.assertIn("zazemlitel", manual_keys)
        self.assertIn("vrezka", manual_keys)


class TestOverrunFactorInInputs(unittest.TestCase):
    def test_overrun_in_inputs_sheet(self):
        wb, _ = _write_and_load(_make_inputs())
        ws = wb[SHEET_INPUTS]
        # B13 = pipe_overrun_factor
        self.assertEqual(ws.cell(13, 2).value, 1.10)


class TestSpecSheetProjection(unittest.TestCase):
    def test_spec_sheet_is_generated_from_vr(self):
        block = BlockRecord(
            handle="f1",
            block_name_raw="СО_Футляр_ПЭ",
            block_name_effective="СО_Футляр_ПЭ",
            layer="!ГСН_Арм",
            position=(0, 0, 0),
            rotation_rad=0,
            scale=(1, 1, 1),
            attributes={"DN": "Футляр DN250x22.7", "LENGTH": "9.36"},
        )
        wb, _ = _write_and_load(_make_inputs(blocks=[block]))
        self.assertIn(SHEET_SPEC, wb.sheetnames)
        ws = wb[SHEET_SPEC]
        formulas = [
            ws.cell(r, 4).value for r in range(1, ws.max_row + 1)
            if isinstance(ws.cell(r, 4).value, str)
        ]
        self.assertTrue(any(v.startswith("='ВР'!D") for v in formulas))
        qty_formulas = [
            ws.cell(r, 9).value for r in range(1, ws.max_row + 1)
            if isinstance(ws.cell(r, 9).value, str)
        ]
        self.assertTrue(any(v.startswith("='ВР'!F") for v in qty_formulas))


if __name__ == "__main__":
    unittest.main()
