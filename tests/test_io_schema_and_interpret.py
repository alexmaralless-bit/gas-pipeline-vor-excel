"""
Tests for pipe_engine.io_schema (JSON loader) and pipe_engine.interpret
(layer/DN parsers + Branch builder).

Fixture: tests/fixtures/demo_extract.v1.json — hand-crafted
extract resembling the actual reference dwg layout (mixed trace layers,
scenery polyline, dynamic block with anonymous raw name, СО_КТ_по_трассе
with DУ instead of DN, etc.).
"""
from __future__ import annotations

import math
import unittest
from pathlib import Path

from pipe_engine.io_schema import (
    SCHEMA_VERSION,
    BlockRecord,
    ExtractDocument,
    PolylineRecord,
)
from pipe_engine.interpret import (
    EXIT_SAND_DEFAULTS,
    TraceLayerKey,
    aggregate_traces,
    build_branches,
    count_consumer_inlet_polylines,
    count_fittings,
    derive_exit_sand_from_extract,
    parse_dn_from_attr,
    parse_trace_layer,
)
from pipe_engine.trench_profile import Section, SoilType


FIXTURE = Path(__file__).resolve().parent / "fixtures" / "demo_extract.v1.json"


# ============================================================================
# Loader (io_schema)
# ============================================================================

class TestIoSchema(unittest.TestCase):
    def test_load_fixture(self):
        doc = ExtractDocument.from_json_file(FIXTURE)
        self.assertEqual(doc.schema_version, SCHEMA_VERSION)
        self.assertEqual(len(doc.polylines), 7)
        self.assertEqual(len(doc.blocks), 5)
        self.assertEqual(doc.block_filter.matched_count, 5)

    def test_polyline_record_shape(self):
        doc = ExtractDocument.from_json_file(FIXTURE)
        p = doc.polylines[0]
        self.assertIsInstance(p, PolylineRecord)
        self.assertEqual(p.handle, "5EA5")
        self.assertEqual(p.layer, "!!ГСН3_Г1_ГП@63")
        self.assertAlmostEqual(p.length, 128.367, places=3)
        self.assertEqual(p.vertex_count, 4)
        self.assertFalse(p.is_closed)

    def test_block_record_shape(self):
        doc = ExtractDocument.from_json_file(FIXTURE)
        b = next(x for x in doc.blocks if x.block_name_effective == "СО_Отвод90_ПЭ+")
        self.assertEqual(b.block_name_raw, "*U37")  # dynamic anonymous → resolved
        self.assertEqual(b.layer, "!ГСН_Арм")
        self.assertEqual(b.attributes["8_МАССА"], "1,22")
        self.assertIn("DN", b.attributes)

    def test_source_meta(self):
        doc = ExtractDocument.from_json_file(FIXTURE)
        self.assertEqual(doc.source.insunits_raw, 4)
        self.assertEqual(doc.source.drawing_units_assumed, "m")

    def test_schema_version_mismatch_rejected(self):
        with self.assertRaises(ValueError):
            ExtractDocument.from_dict({
                "schema_version": "999",
                "source": {}, "polylines": [], "blocks": [], "block_filter": {},
            })


# ============================================================================
# Layer parser
# ============================================================================

class TestParseTraceLayer(unittest.TestCase):
    def test_main_g1_dn110(self):
        k = parse_trace_layer("!!ГСН3_Г1_ГП@110")
        self.assertEqual(k, TraceLayerKey(stage=3, pressure_category="Г1", branch_type=None, dn_mm=110))
        self.assertEqual(k.branch_role, "main")

    def test_consumer_inlets_g1_dn63(self):
        k = parse_trace_layer("!!ГСН3_Г1_ГП_Вводы@63")
        self.assertIsNotNone(k)
        self.assertEqual(k.branch_type, "Вводы")
        self.assertEqual(k.branch_role, "consumer_inlets")
        self.assertEqual(k.dn_mm, 63)

    def test_legacy_g3(self):
        k = parse_trace_layer("!!ГСН2_Г3_ГП@90")
        self.assertEqual(k.stage, 2)
        self.assertEqual(k.pressure_category, "Г3")
        self.assertEqual(k.dn_mm, 90)

    def test_single_exclam_also_matches(self):
        # Some layers have only one `!` — should still match
        k = parse_trace_layer("!ГСН3_Г1_ГП@315")
        self.assertIsNotNone(k)
        self.assertEqual(k.dn_mm, 315)

    def test_project_pe_layer_with_wall_thickness(self):
        k = parse_trace_layer("!ГСН_Г1_ПЭ_@110x10")
        self.assertIsNotNone(k)
        self.assertEqual(k.stage, 1)
        self.assertEqual(k.pressure_category, "Г1")
        self.assertEqual(k.branch_role, "main")
        self.assertEqual(k.dn_mm, 110)

    def test_project_pe_inlet_layer_with_wall_thickness(self):
        k = parse_trace_layer("!ГСН_Г1_ПЭ_Вводы_@63x5.8")
        self.assertIsNotNone(k)
        self.assertEqual(k.stage, 1)
        self.assertEqual(k.branch_type, "Вводы")
        self.assertEqual(k.branch_role, "consumer_inlets")
        self.assertEqual(k.dn_mm, 63)

    def test_non_trace_layer_returns_none(self):
        self.assertIsNone(parse_trace_layer("USLZNСЪЕМКА"))
        self.assertIsNone(parse_trace_layer("!ГСН_Арм"))
        self.assertIsNone(parse_trace_layer("0"))

    def test_missing_dn_returns_none(self):
        # Layer without @DN suffix — not a parseable trace layer
        self.assertIsNone(parse_trace_layer("!!ГСН2_Г3_ГП"))


# ============================================================================
# DN attribute parser
# ============================================================================

class TestParseDn(unittest.TestCase):
    def test_simple_dn(self):
        self.assertEqual(parse_dn_from_attr("Заглушка DN110"), (110, None))

    def test_dn_no_separator(self):
        self.assertEqual(parse_dn_from_attr("DN63"), (63, None))

    def test_transition_cyrillic_х(self):
        # cyrillic х as separator (most common in Russian engineering)
        self.assertEqual(parse_dn_from_attr("Седловой отвод DN110х63"), (110, 63))

    def test_transition_latin_x(self):
        self.assertEqual(parse_dn_from_attr("Переход DN315x225"), (315, 225))

    def test_with_extra_text_around(self):
        self.assertEqual(parse_dn_from_attr("Отвод 90%%d с ЗН DN110"), (110, None))

    def test_ducat_notation(self):
        # "Ду300" alt notation (cyrillic Ду = D Y) — should also work
        self.assertEqual(parse_dn_from_attr("Ду300"), (300, None))

    def test_no_dn_match(self):
        self.assertEqual(parse_dn_from_attr("Кран КШ-150с в ограждении"), (None, None))

    def test_empty_string(self):
        self.assertEqual(parse_dn_from_attr(""), (None, None))


# ============================================================================
# Aggregation
# ============================================================================

class TestAggregateTraces(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.doc = ExtractDocument.from_json_file(FIXTURE)
        cls.aggs = aggregate_traces(cls.doc.polylines)

    def test_scenery_ignored(self):
        # USLZNСЪЕМКА polyline (handle 7001) should NOT appear in aggregates
        layers_in_aggs = {l for a in self.aggs for l in a.source_layers}
        self.assertNotIn("USLZNСЪЕМКА", layers_in_aggs)

    def test_closed_polylines_on_trace_layer_are_ignored(self):
        # Real project pipe layers may contain closed triangle markers on the
        # same layer as the route. They are symbols, not trace length.
        section_doc = ExtractDocument(
            schema_version=SCHEMA_VERSION,
            source=self.doc.source,
            polylines=[
                PolylineRecord(
                    handle="route",
                    layer="!ГСН_Г1_ПЭ_@110x10",
                    length=119.98,
                    start=(0, 0, 0),
                    end=(10, 0, 0),
                    is_closed=False,
                    vertex_count=10,
                ),
                PolylineRecord(
                    handle="triangle",
                    layer="!ГСН_Г1_ПЭ_@110x10",
                    length=6.0,
                    start=(0, 0, 0),
                    end=(0, 0, 0),
                    is_closed=True,
                    vertex_count=3,
                ),
            ],
            blocks=[],
            block_filter=self.doc.block_filter,
        )
        aggs = aggregate_traces(section_doc.polylines)
        self.assertEqual(len(aggs), 1)
        self.assertAlmostEqual(aggs[0].total_length, 119.98, places=2)

    def test_main_g1_dn63_sum(self):
        # Two polylines on !!ГСН3_Г1_ГП@63: 128.367 + 808.07
        main_63 = next(
            a for a in self.aggs
            if a.key.branch_role == "main" and a.key.dn_mm == 63
        )
        self.assertEqual(main_63.polyline_count, 2)
        self.assertAlmostEqual(main_63.total_length, 128.367 + 808.07, places=3)

    def test_inlets_g1_dn63_sum(self):
        # Two polylines on !!ГСН3_Г1_ГП_Вводы@63: 7.5 + 8.1
        inlets_63 = next(
            a for a in self.aggs
            if a.key.branch_role == "consumer_inlets" and a.key.dn_mm == 63
        )
        self.assertEqual(inlets_63.polyline_count, 2)
        self.assertAlmostEqual(inlets_63.total_length, 15.54, places=3)

    def test_dn_buckets_present(self):
        dns_for_main = sorted(
            a.key.dn_mm for a in self.aggs if a.key.branch_role == "main"
        )
        self.assertEqual(dns_for_main, [63, 110, 160])


# ============================================================================
# Branch builder integration
# ============================================================================

class TestBuildBranches(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.doc = ExtractDocument.from_json_file(FIXTURE)
        cls.section = Section(name="demo", soil=SoilType.SAND, season="winter")
        cls.branches = build_branches(cls.doc, default_section=cls.section)

    def test_two_branches(self):
        # Expect one "main" branch and one "consumer_inlets" branch
        roles = sorted({b.id for b in self.branches})
        self.assertEqual(len(self.branches), 2)
        # IDs are deterministic: e.g. "main_г1_stage3" and "consumer_inlets_г1_stage3"
        self.assertTrue(any("main" in r for r in roles))
        self.assertTrue(any("consumer_inlets" in r for r in roles))

    def test_main_has_three_pipe_rows(self):
        main = next(b for b in self.branches if b.id.startswith("main_"))
        dns = sorted(p.dn_mm for p in main.pipes)
        self.assertEqual(dns, [63, 110, 160])

    def test_inlets_single_pipe_row(self):
        inlets = next(b for b in self.branches if b.id.startswith("consumer_inlets"))
        self.assertEqual(len(inlets.pipes), 1)
        self.assertEqual(inlets.pipes[0].dn_mm, 63)
        self.assertAlmostEqual(inlets.pipes[0].length_grp1, 15.54, places=3)

    def test_branches_share_default_section(self):
        for b in self.branches:
            self.assertEqual(b.section.name, "demo")
            self.assertEqual(b.section.season, "winter")


# ============================================================================
# Fitting counts
# ============================================================================

class TestFittingCounts(unittest.TestCase):
    def test_counts_match_fixture(self):
        doc = ExtractDocument.from_json_file(FIXTURE)
        counts = {c.block_name: c.count for c in count_fittings(doc.blocks)}
        self.assertEqual(counts["СО_Отвод90_ПЭ+"], 1)
        self.assertEqual(counts["СО_Переход_ПЭ"], 1)
        self.assertEqual(counts["СО_КТ_по_трассе"], 1)
        self.assertEqual(counts["СО_ЗУ_кран"], 1)
        self.assertEqual(counts["СО_Седло_Отвод_ПЭ"], 1)


class TestConsumerInletPolylineCount(unittest.TestCase):
    """Узел А count = polylines on _Вводы@* layers (engineering rule)."""

    def test_fixture_has_two_inlet_polylines(self):
        doc = ExtractDocument.from_json_file(FIXTURE)
        # Fixture has two polylines on !!ГСН3_Г1_ГП_Вводы@63
        n = count_consumer_inlet_polylines(doc)
        self.assertEqual(n, 2)


class TestDeriveExitSand(unittest.TestCase):
    """Canonical exit-sand derivation from dwg extract."""

    @classmethod
    def setUpClass(cls):
        cls.doc = ExtractDocument.from_json_file(FIXTURE)

    def test_no_zaglushka_in_exit_sand(self):
        # СО_Загл_ПЭ is a buried pipe end-cap, NOT a surface exit.
        # Even if fixture has Заглушка blocks, derive_exit_sand must skip them.
        nodes = derive_exit_sand_from_extract(self.doc)
        names = [n.name for n in nodes]
        for n in names:
            self.assertNotIn("Загл", n,
                f"Exit sand should not include Заглушка blocks; got {names}")

    def test_zu_kran_kratность_4(self):
        nodes = derive_exit_sand_from_extract(self.doc)
        zu = next((n for n in nodes if "ЗУ надземно" in n.name), None)
        self.assertIsNotNone(zu, "ЗУ надземно node missing")
        self.assertEqual(zu.sand_units, 4)
        # Fixture has exactly 1 СО_ЗУ_кран
        self.assertEqual(zu.count, 1)

    def test_uzel_a_kratность_2_from_polyline_count(self):
        # Fixture has 2 inlet polylines → 2 Узлы А
        nodes = derive_exit_sand_from_extract(self.doc)
        ua = next((n for n in nodes if "Узел А" in n.name), None)
        self.assertIsNotNone(ua, "Узел А node missing")
        self.assertEqual(ua.sand_units, 2)
        self.assertEqual(ua.count, 2)

    def test_kt_traссе_kratность_1(self):
        nodes = derive_exit_sand_from_extract(self.doc)
        kt = next((n for n in nodes if n.name == "КТ по трассе"), None)
        self.assertIsNotNone(kt)
        self.assertEqual(kt.sand_units, 1)

    def test_no_internal_fittings_in_exit_sand(self):
        # СО_Седло_Отвод, СО_Переход, СО_Тройник, СО_Муфта, СО_Отвод — all
        # internal pipe fittings, fully buried, should NOT contribute.
        nodes = derive_exit_sand_from_extract(self.doc)
        names = [n.name for n in nodes]
        bad_patterns = ["Седло", "Переход", "Тройник", "Муфта", "Отвод90"]
        for needle in bad_patterns:
            for n in names:
                self.assertNotIn(needle, n,
                    f"{needle!r} should not be in exit_sand list; got {names}")

    def test_hand_trench_count_passes_through(self):
        nodes = derive_exit_sand_from_extract(self.doc, hand_trench_count=92)
        ht = next((n for n in nodes if "Пересеч" in n.name), None)
        self.assertIsNotNone(ht)
        self.assertEqual(ht.count, 92)
        self.assertEqual(ht.sand_units, 1)

    def test_gnb_count_passes_through(self):
        nodes = derive_exit_sand_from_extract(self.doc, gnb_pit_count=44)
        gnb = next((n for n in nodes if "ГНБ" in n.name), None)
        self.assertIsNotNone(gnb)
        self.assertEqual(gnb.count, 44)

    def test_zero_counts_omit_node(self):
        # If a category has 0 count and 0 manual param, it's omitted entirely
        nodes = derive_exit_sand_from_extract(self.doc, hand_trench_count=0, gnb_pit_count=0)
        names = [n.name for n in nodes]
        self.assertNotIn("Пересеч инж комм в футляре", names)
        self.assertNotIn("Прокладка в футляре ГНБ", names)


if __name__ == "__main__":
    unittest.main()
