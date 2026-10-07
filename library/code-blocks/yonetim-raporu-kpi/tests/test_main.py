import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
VERI = """Tarih;Fatura No;Müşteri;Ürün;Kategori;Kanal;Adet;Net Tutar;Maliyet;Tür
10.09.2025;F1;M1;U1;K1;Bayi;10;1000;600;Satış
15.08.2026;F2;M1;U1;K1;Bayi;10;1200;700;Satış
05.09.2026;F3;M1;U1;K1;Bayi;10;1500;900;Satış
06.09.2026;F4;M2;U2;K2;Online;5;500;200;Satış
20.09.2026;I5;M2;U2;K2;Online;1;100;40;İade
"""


class KpiTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "s.csv").write_text(VERI, encoding="utf-8")
            cls.s = main.calistir(t / "s.csv", t / "y.xlsx", (2026, 9))
        cls.k = cls.s["kpi"]

    def test_rapor_ayi(self):
        r = self.k["Rapor ayı"]
        self.assertEqual(r["satis"], D(1900))                         # 1500 + 500 − 100 (iade ters çevrilir)
        self.assertEqual(r["brut_kar"], D(1900 - (900 + 200 - 40)))   # 840
        self.assertEqual(r["fatura"], 2)                               # iade fatura sayısına girmez
        self.assertEqual(r["ort_fatura"], D(950))
        self.assertEqual((r["aktif_musteri"], r["yeni_musteri"]), (2, 1))   # M2 ilk kez bu ay
        self.assertEqual(r["ilk5_pay"], 1)

    def test_karsilastirmalar(self):
        self.assertEqual(self.k["Önceki ay"]["satis"], D(1200))
        self.assertEqual(self.k["Geçen yıl aynı ay"]["satis"], D(1000))
        self.assertEqual(self.k["YTD"]["satis"], D(3100))
        self.assertAlmostEqual(main.degisim(self.k["Rapor ayı"]["satis"], self.k["Geçen yıl aynı ay"]["satis"]), 0.9)

    def test_kirilim_ve_trend(self):
        self.assertEqual(self.s["kirilim"]["kanal"]["Online"][0], D(400))
        self.assertEqual(len(self.s["trend"]), 13)
        self.assertEqual(list(self.s["trend"])[0], (2025, 9))


class OrnekTesti(unittest.TestCase):
    def test_mizanli(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "satislar.csv", Path(tmp) / "y.xlsx", None, ORNEK / "mizan_eylul_2026.csv")
        self.assertEqual(s["ay"], (2026, 9))
        self.assertGreater(s["kpi"]["Rapor ayı"]["satis"], 0)
        self.assertEqual(dict(s["fin"]["satirlar"])["Dönem net kârı"], D(2925000))

if __name__ == "__main__":
    unittest.main()
