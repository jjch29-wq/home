import tempfile
import unittest
from pathlib import Path

from core import Point, Project, export_dxf, export_pdf, load_project, save_project, snap_iso


class CoreTests(unittest.TestCase):
    def test_snap(self):
        x,y=snap_iso((0,0),(10,2)); self.assertAlmostEqual(y/x,0,places=6)
        x,y=snap_iso((0,0),(10,7)); self.assertAlmostEqual(y/x,3**.5/3,places=6)

    def test_round_trip_and_exports(self):
        p=Project(points=[Point(0,0),Point(100,0,1200,"ELBOW_90"),Point(150,86.6,800)])
        with tempfile.TemporaryDirectory() as d:
            root=Path(d); save_project(p,root/"a.json"); q=load_project(root/"a.json")
            self.assertEqual(q.points[1].actual_length,1200)
            export_dxf(q,root/"a.dxf"); export_pdf(q,root/"a.pdf")
            self.assertIn("LINE",(root/"a.dxf").read_text())
            self.assertTrue((root/"a.pdf").read_bytes().startswith(b"%PDF"))


if __name__ == "__main__": unittest.main()
