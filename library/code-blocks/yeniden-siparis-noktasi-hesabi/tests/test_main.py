import math
import sys
import tempfile
import unittest
from pathlib import Path
from statistics import NormalDist

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
Z95 = NormalDist().inv_cdf(0.95)          # 1,6449


class FormulTesti(unittest.TestCase):
    def test_emniyet_stogu_sabit_tedarik(self):
        # d = 20, σd = 4, L = 9 gün, %95: SS = 1,645 × 4 × √9 = 19,74
        self.assertAlmostEqual(main.emniyet_stogu(Z95, 20, 4, 9, 0), Z95 * 12)
        self.assertAlmostEqual(main.emniyet_stogu(Z95, 20, 4, 9, 0), 19.74, places=2)

    def test_emniyet_stogu_degisken_tedarik(self):
        # d = 50, σd = 10, L = 6, σL = 1: √(6 × 100 + 2500 × 1) = √3100
        self.assertAlmostEqual(main.emniyet_stogu(Z95, 50, 10, 6, 1), Z95 * math.sqrt(3100))

    def test_esm(self):
        # D = 10.000, S = 50 TL, h = %20 × 10 TL = 2 TL → √500.000 = 707,1
        self.assertAlmostEqual(main.esm(10000, 50, 0.20, 10), math.sqrt(500000))
        self.assertIsNone(main.esm(10000, 0, 0.2, 10))

    def test_yuvarlama(self):
        self.assertEqual(main.yuvarla_kat(2163.7, 1000, 25), 2175)
        self.assertEqual(main.yuvarla_kat(10.2, 25, 25), 25)
        self.assertEqual(main.yuvarla_kat(10.2, 0, 0), 11)


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "kalemler.csv", Path(tmp) / "y.xlsx", ORNEK / "aylik_tuketim.csv", 30, 95, 750, 0.25)
        cls.k = {x["kod"]: x for x in cls.s["kalemler"]}

    def test_ders_kitabi_ornekleri(self):
        self.assertAlmostEqual(self.k["X-500"]["ysn"], 180 + Z95 * 12)
        self.assertAlmostEqual(self.k["X-501"]["ysn"], 300 + Z95 * math.sqrt(3100))

    def test_gecmisten_talep(self):
        hm = self.k["HM-001"]
        self.assertAlmostEqual(hm["d"], 48100 / 12 / 30)                  # aylık ortalama / 30
        self.assertEqual(hm["kaynak"], "geçmiş (12 dönem)")
        self.assertAlmostEqual(hm["hizmet"], 0.98)

    def test_siparis_onerisi(self):
        hm = self.k["HM-001"]
        self.assertEqual(hm["pozisyon"], 1300)                             # 1500 − 200 ayrılmış
        self.assertTrue(hm["acil"])                                        # 1300 / 133,6 = 9,7 gün < 10 gün tedarik
        self.assertEqual(hm["oneri"] % 25, 0)                              # çuval katı
        self.assertGreaterEqual(hm["oneri"], 1000)                         # en az sipariş
        self.assertEqual(self.k["YM-010"]["oneri"], 0)                     # pozisyon YSN'nin üstünde

    def test_gecersiz_hizmet(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "k.csv").write_text("Stok Kodu;Ortalama Günlük Talep;Günlük Talep Std;Tedarik Süresi (gün);Hizmet Düzeyi\nA;1;1;1;100\n",
                                     encoding="utf-8")
            with self.assertRaises(SystemExit):
                main.calistir(t / "k.csv", t / "y.xlsx")


if __name__ == "__main__":
    unittest.main()
