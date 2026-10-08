import sys
import tempfile
import unittest
from collections import OrderedDict
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
ASORTI = OrderedDict([("S", 1), ("M", 2), ("L", 2), ("XL", 1)])


class PlanTesti(unittest.TestCase):
    def test_solid_ve_karisik(self):
        g = OrderedDict({("P", "M", "R"): OrderedDict([("S", 50), ("M", 30)])})
        s = main.koli_plani(g, "solid", 24, None)
        self.assertEqual([(x["tur"], x["icerik"], x["koli"]) for x in s],
                         [("Solid", {"S": 24}, 2), ("Solid", {"M": 24}, 1), ("Karışık", {"S": 2, "M": 6}, 1)])

    def test_asorti(self):
        g = OrderedDict({("P", "M", "R"): OrderedDict([("S", 150), ("M", 250), ("L", 250), ("XL", 100)])})
        s = main.koli_plani(g, "asorti", 24, ASORTI)
        self.assertEqual((s[0]["tur"], s[0]["icerik"], s[0]["koli"]), ("Asorti", {"S": 4, "M": 8, "L": 8, "XL": 4}, 25))
        self.assertEqual(sum(sum(x["icerik"].values()) * x["koli"] for x in s), 750)


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as t:
            cls.s = main.calistir(ORNEK / "sevk.csv", Path(t) / "a.xlsx", "asorti", 24, ASORTI, 0.25, 0.8, (60, 40, 40))
        cls.r = cls.s["satirlar"]

    def test_numaralar_surekli(self):
        self.assertEqual(self.r[0]["bas"], 1)
        for a, b in zip(self.r, self.r[1:]):
            self.assertEqual(b["bas"], a["son"] + 1)
        self.assertEqual(self.r[-1]["son"], 159)

    def test_agirlik_ve_hacim(self):
        ilk = self.r[0]
        self.assertAlmostEqual(ilk["net_koli"], 24 * 0.22)                   # dosyadaki adet ağırlığı
        self.assertAlmostEqual(ilk["brut_koli"], 24 * 0.22 + 0.8)
        self.assertAlmostEqual(ilk["hacim"], 75 * 0.096)

    def test_tum_adetler_paketlendi(self):
        self.assertEqual(sum(x["toplam"] for x in self.r), 3770)

    def test_hatali_asorti(self):
        with tempfile.TemporaryDirectory() as t:
            with self.assertRaises(SystemExit):
                main.calistir(ORNEK / "sevk.csv", Path(t) / "a.xlsx", "asorti", 25, ASORTI)


if __name__ == "__main__":
    unittest.main()
