"""Tests for pipe_engine.node_library — template loader + instantiation."""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from pipe_engine.node_library import (
    InstantiatedNode,
    NodeComponent,
    NodeLibrary,
    NodeTemplate,
    NodeVariant,
)


class TestDefaultLibrary(unittest.TestCase):
    """Templates shipped in pipe_engine/data/node_templates/."""

    @classmethod
    def setUpClass(cls):
        cls.lib = NodeLibrary.load_default()

    def test_at_least_two_templates(self):
        self.assertGreaterEqual(len(self.lib), 2)

    def test_kt_template_present(self):
        t = self.lib.find_by_id("kt_po_trasse")
        self.assertIsNotNone(t)
        self.assertEqual(t.detected_by_block, "СО_КТ_по_трассе")
        self.assertEqual(len(t.components), 3)

    def test_zu_template_present(self):
        t = self.lib.find_by_id("zu_kran_nadzemno")
        self.assertIsNotNone(t)
        self.assertEqual(t.detected_by_block, "СО_ЗУ_кран")

    def test_find_by_block_works(self):
        self.assertIsNotNone(self.lib.find_by_block("СО_КТ_по_трассе"))
        self.assertIsNotNone(self.lib.find_by_block("СО_ЗУ_кран"))
        # truly unknown block name — not in any shipped template
        self.assertIsNone(self.lib.find_by_block("СО_НЕИЗВЕСТНЫЙ_БЛОК"))

    def test_real_project_block_aliases_work(self):
        self.assertEqual(self.lib.find_by_block("СО_Переход_ПЭ-Сталь").id, "perehod_pe")
        self.assertEqual(self.lib.find_by_block("СО_Отвод_90_ПЭ_литой").id, "otvod90_pe")


class TestVariantPicking(unittest.TestCase):
    """pick_variant logic — exact match, closest-down fallback, edge cases."""

    @classmethod
    def setUpClass(cls):
        cls.kt = NodeLibrary.load_default().find_by_id("kt_po_trasse")

    def test_exact_match(self):
        v = self.kt.pick_variant(100)
        self.assertEqual(v.host_dn, 100)
        self.assertEqual(v.suffix, "03")
        self.assertEqual(v.L_mm, 400)

    def test_closest_smaller(self):
        # 63 doesn't exist in КТ table — closest smaller is 50
        v = self.kt.pick_variant(63)
        self.assertEqual(v.host_dn, 50)

    def test_above_max_returns_first(self):
        # DN 99999 — no exact, all smaller works, returns max smaller (600)
        v = self.kt.pick_variant(99999)
        self.assertEqual(v.host_dn, 600)

    def test_below_min_returns_first(self):
        # DN 25 — no smaller available, returns first variant
        v = self.kt.pick_variant(25)
        self.assertEqual(v.host_dn, 50)

    def test_none_dn_returns_first(self):
        v = self.kt.pick_variant(None)
        self.assertEqual(v.host_dn, 50)

    def test_template_without_variants(self):
        t = NodeTemplate(id="x", name="x", detected_by_block="X", components=[])
        self.assertIsNone(t.pick_variant(100))


class TestInstantiate(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.lib = NodeLibrary.load_default()

    def test_instantiate_known(self):
        inst = self.lib.instantiate("СО_КТ_по_трассе", block_handle="ABC", host_dn_mm=100)
        self.assertIsNotNone(inst)
        self.assertEqual(inst.template_id, "kt_po_trasse")
        self.assertEqual(inst.block_handle, "ABC")
        self.assertEqual(inst.host_dn_mm, 100)
        self.assertEqual(inst.variant.suffix, "03")
        self.assertEqual(len(inst.components), 3)

    def test_instantiate_unknown_block(self):
        inst = self.lib.instantiate("СО_НЕИЗВЕСТНЫЙ_БЛОК", host_dn_mm=100)
        self.assertIsNone(inst)


class TestLoaderRobustness(unittest.TestCase):
    """Loader error cases + JSON/YAML format support."""

    def test_load_missing_id_raises(self):
        with tempfile.TemporaryDirectory() as td:
            bad = Path(td) / "bad.json"
            bad.write_text(json.dumps({"name": "x", "detected_by_block": "X",
                                       "components": []}), encoding="utf-8")
            with self.assertRaises(ValueError):
                NodeLibrary.load_from_directory(td)

    def test_load_empty_directory(self):
        with tempfile.TemporaryDirectory() as td:
            lib = NodeLibrary.load_from_directory(td)
            self.assertEqual(len(lib), 0)

    def test_load_nonexistent_directory(self):
        with self.assertRaises(FileNotFoundError):
            NodeLibrary.load_from_directory("/nonexistent")

    def test_components_default_kind_material(self):
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "t.json"
            p.write_text(json.dumps({
                "id": "t", "name": "T", "detected_by_block": "X",
                "components": [{"tag": "X", "qty": 1}],
            }), encoding="utf-8")
            lib = NodeLibrary.load_from_directory(td)
            self.assertEqual(lib.templates[0].components[0].kind, "material")
            self.assertEqual(lib.templates[0].components[0].unit, "шт")


if __name__ == "__main__":
    unittest.main()
