import unittest
from types import SimpleNamespace
import pandas as pd
from test_central_inspection_batch import Field
from site_apps.central.src.method_catalog import classify, material_candidates
from site_apps.central.src.aux_equipment import equipment_catalog
from site_apps.central.src.material_selection import choose_daily_material, resolve_daily_material


class MethodCatalogTests(unittest.TestCase):
    def app(self):
        rows = [
            (1, 'MT약품', '백색페인트', 'NDT_CHEM'),
            (2, 'PT약품', '침투제', 'NDT_CHEM'),
            (3, 'Carestream T200', '12인치', 'FILM'),
            (4, 'PAUT장비', 'MX2', 'UT장비'),
            (5, 'COBRA Manual Scanner', 'COBRA', 'UT장비'),
            (6, 'YOKE AC', 'MY-2', 'MT장비'),
            (7, 'Black Light', 'UV LED', '계측장비'),
            (8, '확인 필요', '모델X', ''),
            (9, '초음파 탐상기', 'Sitescan', 'UT장비'),
        ]
        app = SimpleNamespace(cb_daily_test_method=Field('MT'), cb_daily_material=Field(''),
                              equipments=[], aux_equipments=[])
        app.materials_df = pd.DataFrame([dict(MaterialID=i, 품목명=n, 모델명=m, 품목군코드=c,
                                             SN=f'SN{i}', Active=1) for i,n,m,c in rows])
        app.get_material_display_name = lambda mid: next(f'{n} ({m})' for i,n,m,c in rows if i == int(mid))
        return app

    def test_materials_follow_method_without_mutating_selected_id_or_stock(self):
        app = self.app()
        before = app.materials_df.copy(deep=True)
        choose_daily_material(app, 1)
        self.assertEqual(material_candidates(app), ['MT약품 (백색페인트)'])
        app.cb_daily_test_method.set('PT')
        self.assertEqual(material_candidates(app), ['PT약품 (침투제)'])
        self.assertEqual(resolve_daily_material(app), 1)  # old record stays intact
        pd.testing.assert_frame_equal(before, app.materials_df)

    def test_roles_and_unknown_opt_in_keep_existing_selection(self):
        app = self.app()
        app.cb_daily_test_method.set('PAUT')
        main, aliases = equipment_catalog(app, 'main', 'YOKE AC MY-2 (SN6)')
        aux, _ = equipment_catalog(app, 'aux')
        self.assertTrue(any('MX2' in s for s in main))
        self.assertFalse(any('COBRA' in s or 'YOKE' in s for s in main))
        self.assertTrue(any('COBRA' in s for s in aux))
        self.assertIn('YOKE AC MY-2 (SN6)', aliases)
        self.assertFalse(any('모델X' in s for s in main))
        extra, _ = equipment_catalog(app, 'main', include_unknown=True)
        self.assertTrue(any('모델X' in s for s in extra))
        app.cb_daily_test_method.set('UT')
        ut, _ = equipment_catalog(app, 'main')
        self.assertTrue(any('Sitescan' in s for s in ut))
        self.assertFalse(any('MX2' in s for s in ut))

    def test_shared_instruments_and_specific_probe_rules(self):
        self.assertEqual(classify('Black Light', '계측장비'), ({'MT', 'PT'}, 'aux'))
        self.assertEqual(classify('PAUT PROBE(A2)', 'UT자재'), ({'PAUT'}, 'aux'))
        self.assertEqual(classify('UT PROBE', 'UT자재'), ({'UT'}, 'aux'))
        self.assertEqual(classify('TEST BLOCK', 'UT자재'), ({'UT', 'PAUT'}, 'aux'))
        self.assertEqual(classify('Survey Meter', '안전장비'), ({'RT'}, 'aux'))


if __name__ == '__main__':
    unittest.main()
