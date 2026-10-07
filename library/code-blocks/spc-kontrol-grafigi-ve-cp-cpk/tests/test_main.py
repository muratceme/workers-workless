import statistics
import sys
import tempfile
import unittest
from collections import OrderedDict
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402


class SinirTesti(unittest.TestCase):
    def test_elle_hesaplanan_sinirlar(self):
        # n=4, 3 alt grup: ortalamalar 10, 11, 12 → X̄̄ = 11; aralıklar 2, 2, 2 → R̄ = 2
        g = OrderedDict([("1", [9, 10, 11, 10]), ("2", [10, 11, 12, 11]), ("3", [11, 12, 13, 12])])
        s = main.hesapla(g, 5, 17)
        self.assertAlmostEqual(s.x_ort, 11)
        self.assertAlmostEqual(s.r_ort, 2)
        self.assertAlmostEqual(s.x_ust, 11 + 0.729 * 2)
        self.assertAlmostEqual(s.x_alt, 11 - 0.729 * 2)
        self.assertAlmostEqual(s.r_ust, 2.282 * 2)
        self.assertAlmostEqual(s.r_alt, 0)
        sigma = 2 / 2.059
        self.assertAlmostEqual(s.yeterlilik["Cp"], 12 / (6 * sigma))
        self.assertAlmostEqual(s.yeterlilik["Cpk"], 6 / (3 * sigma))
        tum = [x for v in g.values() for x in v]
        self.assertAlmostEqual(s.yeterlilik["Pp"], 12 / (6 * statistics.stdev(tum)))
        self.assertTrue(any("en az 20" in u for u in s.uyarilar))

    def test_tek_tarafli_sartname(self):
        g = OrderedDict((str(i), [10.0, 10.2, 9.8]) for i in range(25))
        s = main.hesapla(g, None, 11)
        self.assertNotIn("Cp", s.yeterlilik)
        self.assertIn("Cpk", s.yeterlilik)


class NelsonTesti(unittest.TestCase):
    def test_kurallar(self):
        self.assertIn("K1: 3σ dışında", main.nelson([0, 0, 3.5, 0], 0, 1)[2])
        ihlal = main.nelson([0.5] * 9, 0, 1)
        self.assertIn("K2: 9 nokta merkezin aynı tarafında", ihlal[8])
        ihlal = main.nelson([1, 2, 3, 4, 5, 6], 3.5, 10)
        self.assertIn("K3: 6 nokta sürekli artış/azalış", ihlal[5])
        ihlal = main.nelson([2.5, 0, 2.5], 0, 1)
        self.assertIn("K5: 3 noktanın 2'si 2σ dışında (aynı taraf)", ihlal[2])
        self.assertEqual(main.nelson([0.1, -0.2, 0.3, -0.1], 0, 1), {})


class BicimTesti(unittest.TestCase):
    def test_uzun_bicim(self):
        satirlar = [["Alt Grup", "Ölçüm"]] + [[str(g), v] for g in range(3) for v in (1, 2, 3)]
        g = main.alt_gruplari_cikar(satirlar)
        self.assertEqual(list(g), ["0", "1", "2"])
        self.assertEqual(g["1"], [1.0, 2.0, 3.0])


class UctanUcaTest(unittest.TestCase):
    def test_ornek(self):
        g = main.alt_gruplari_cikar(main.oku(KLASOR / "ornek_veri" / "mil_capi.csv"))
        s = main.hesapla(g, 24.95, 25.05)
        self.assertEqual((s.n, len(s.alt_grup_adlari)), (5, 25))
        self.assertTrue(s.ihlaller or s.r_ihlaller)           # örnek veride kayma ve aykırı değer var
        self.assertIn(12, s.r_ihlaller)                       # 13. alt gruptaki 25,075 ölçümü
        with tempfile.TemporaryDirectory() as tmp:
            main.rapor_yaz(s, Path(tmp) / "s.xlsx", "Mil çapı", 24.95, 25.05)


if __name__ == "__main__":
    unittest.main()
