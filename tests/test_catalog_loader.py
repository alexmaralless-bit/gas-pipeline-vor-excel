"""Tests for SQLite-backed reference data loaded from pipe_catalog.db."""
from __future__ import annotations

import unittest

from pipe_engine.catalog_loader import load_catalog


class TestCatalogReferenceTables(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.catalog = load_catalog()

    def test_dn_translation_loaded(self):
        by_outer = {r.diameter_outer: r.dn for r in self.catalog.dn_translations}
        self.assertEqual(by_outer[108], 100)
        self.assertEqual(by_outer[273], 250)

    def test_climate_reference_loaded(self):
        maykop = next(
            row for row in self.catalog.climate_sp131
            if row.region == "Республика Адыгея" and row.locality == "Майкоп"
        )
        self.assertEqual(maykop.cold_5day_092_c, -19.0)
        self.assertEqual(maykop.duration_below_8_days, 148)
        self.assertEqual(maykop.avg_temp_below_8_c, 2.3)
        self.assertEqual(maykop.avg_wind_speed_below_8_m_s, 3.0)

    def test_work_catalog_loaded(self):
        berm = self.catalog.work_by_key("earth.excess_soil_berm")
        self.assertIsNotNone(berm)
        self.assertEqual(berm.name_for_vor, "Устройство валика из вытесненного грунта")
        self.assertEqual(berm.norm_source, "ФСНБ-2022/ГЭСН/ФЕР pending")

        open_laying = self.catalog.work_by_key("pipe.pe_open_laying")
        self.assertIsNotNone(open_laying)
        self.assertEqual(open_laying.unit, "м")


if __name__ == "__main__":
    unittest.main()
