import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class YardimciTesti(unittest.TestCase):
    def test_sayi_bicimleri(self):
        self.assertEqual(main.sayi("1.234,56"), 1234.56)
        self.assertEqual(main.sayi("1,234.56"), 1234.56)
        self.assertEqual(main.sayi("1,234"), 1234.0)
        self.assertEqual(main.sayi("12,5"), 12.5)
        self.assertEqual(main.sayi("₺2.500"), 2500.0)
        self.assertEqual(main.sayi("--"), 0.0)

    def test_olcutler(self):
        v = main.olcutler({"gosterim": 10000, "tiklama": 200, "harcama": 1000, "donusum": 10, "donusum_degeri": 5000})
        self.assertEqual((v["ctr"], v["cpc"], v["cpm"], v["cvr"], v["cpa"], v["roas"]), (0.02, 5.0, 100.0, 0.05, 100.0, 5.0))


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as t:
            cls.s = main.calistir([ORNEK / "google_ads_kampanya.csv", ORNEK / "meta_ads_kampanya.csv"], Path(t) / "a.xlsx",
                                  hedef_roas=4, hedef_cpa=150, butce_yolu=ORNEK / "butceler.csv")

    def test_sutun_tanima(self):
        g, m = self.s["tanima"]["google_ads_kampanya.csv"], self.s["tanima"]["meta_ads_kampanya.csv"]
        self.assertEqual((g["harcama"], g["donusum"], g["donusum_degeri"]), ("Maliyet", "Dönüşümler", "Dönüşüm değeri"))
        self.assertEqual((m["harcama"], m["tiklama"], m["donusum"]), ("Amount spent (TRY)", "Link clicks", "Purchases"))

    def test_platform_ve_toplam(self):
        self.assertEqual(set(self.s["platform"]), {"Google", "Meta"})
        self.assertEqual(len(self.s["kampanya"]), 7)                              # "Toplam" satırı atlanır
        t = self.s["toplam"]
        self.assertAlmostEqual(t["harcama"], sum(v["harcama"] for v in self.s["platform"].values()))
        self.assertAlmostEqual(t["roas"], t["donusum_degeri"] / t["harcama"])
        self.assertEqual(len(self.s["hafta"]), 5)

    def test_isaretler(self):
        notlar = {k[1]: v["not"] for k, v in self.s["kampanya"].items()}
        self.assertIn("dönüşüm yok", notlar["Görüntülü - Yeniden Pazarlama"])
        self.assertIn("artırılabilir", notlar["Arama - Marka"])
        self.assertIn("ROAS hedefin çok altında", notlar["Prospecting - Lookalike"])

    def test_butce_temposu(self):
        b = {x["kampanya"]: x for x in self.s["butce"]}
        self.assertAlmostEqual(b["Prospecting - Lookalike"]["beklenen"], 45000 * 30 / 61)   # 01.09–31.10, 30 gün geçti
        self.assertAlmostEqual(b["Performance Max"]["beklenen"], 15000)

    def test_esleme(self):
        with tempfile.TemporaryDirectory() as t:
            y = Path(t) / "x.csv"
            y.write_text("Ad Set;Spend TL;Views\nK1;100;1000\n", encoding="utf-8")
            with self.assertRaises(SystemExit):
                main.calistir([y], Path(t) / "a.xlsx")
            s = main.calistir([y], Path(t) / "a.xlsx", esleme={"kampanya": "Ad Set", "harcama": "Spend TL", "gosterim": "Views"})
        self.assertEqual(s["toplam"]["harcama"], 100)


if __name__ == "__main__":
    unittest.main()
