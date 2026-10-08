import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class YardimciTesti(unittest.TestCase):
    def test_is_gunu(self):
        self.assertEqual(main.is_gunu_say(date(2026, 10, 1), date(2026, 10, 31)), 22)
        self.assertEqual(main.is_gunu_say(date(2026, 10, 1), date(2026, 10, 15)), 11)

    def test_prim(self):
        k = [(90, 0.5), (100, 1.0), (110, 1.5)]
        self.assertEqual([main.prim_carpani(x, k) for x in (0.85, 0.95, 1.0, 1.2)], [0, 0.5, 1.0, 1.5])

    def test_ay_anahtari(self):
        for x in ("2026-10", "10.2026", "Ekim 2026", "202610", date(2026, 10, 3)):
            self.assertEqual(main.ay_anahtari(x), "2026-10")


class ElleHesapTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as t:
            t = Path(t)
            (t / "s.csv").write_text("Tarih;Temsilci;Net Tutar;Tür\n01.10.2026;A;1000;Satış\n15.10.2026;A;1200;Satış\n"
                                     "16.10.2026;A;5000;Satış\n10.10.2026;B;300;İade\n03.10.2025;A;1000;Satış\n20.10.2025;A;9000;Satış\n"
                                     "05.01.2026;A;4000;Satış\n", encoding="utf-8")
            (t / "h.csv").write_text("Temsilci;2026-01;2026-10\nA;5000;4400\nB;0;1000\n", encoding="utf-8")
            cls.s = main.calistir(t / "s.csv", t / "h.csv", t / "r.xlsx", "2026-10", date(2026, 10, 15))
        cls.g = next(iter(cls.s["tablolar"]["Genel"][1].values()))
        cls.a = cls.s["tablolar"]["Temsilci"][1][("A",)]

    def test_ay_ve_iade(self):
        self.assertEqual(self.g["ay"], 1000 + 1200 - 300)                 # 16.10 kesimden sonra; iade düşülür
        self.assertEqual(self.a["ay"], 2200)
        self.assertEqual(self.a["hedef_ay"], 4400)

    def test_run_rate_ve_gereken(self):
        self.assertAlmostEqual(self.a["tahmin"], 2200 / 11 * 22)
        self.assertAlmostEqual(self.a["gereken_gunluk"], (4400 - 2200) / 11)
        self.assertAlmostEqual(self.a["beklenen"], 4400 * 11 / 22)

    def test_gy_ve_ytd(self):
        self.assertEqual(self.a["gy_ay"], 1000)                            # 20.10.2025 GY kesiminden (15.10.2025) sonra
        self.assertEqual(self.a["ytd"], 2200 + 4000)
        self.assertEqual(self.a["hedef_ytd"], 5000 + 4400)


class OrnekTest(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as t:
            s = main.calistir(ORNEK / "satislar.csv", ORNEK / "hedefler.csv", Path(t) / "r.xlsx", "2026-09", date(2026, 10, 15),
                              kademeler=[(90, 0.5), (100, 1.0), (110, 1.5)])
        self.assertFalse(s["ay_ici"])
        self.assertEqual(set(s["tablolar"]), {"Genel", "Temsilci", "Bölge", "Temsilci × Ürün Grubu"})
        self.assertTrue(all(v["prim_carpani"] is not None for v in s["tablolar"]["Temsilci"][1].values()))


if __name__ == "__main__":
    unittest.main()
