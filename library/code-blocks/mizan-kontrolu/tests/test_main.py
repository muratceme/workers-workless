import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class DogaTesti(unittest.TestCase):
    def test_doga(self):
        beklenen = {"100": "B", "257": "A", "103": "A", "129": "A", "320": "A", "322": "B", "371": "B", "500": "A",
                    "501": "B", "580": "B", "591": "B", "600": "A", "611": "B", "632": "B", "642": "A", "690": None,
                    "691": "B", "770": "B", "771": "A", "772": None, "798": "A", "790": "B", "900": None}
        self.assertEqual({k: main.doga(k) for k in beklenen}, beklenen)

    def test_ust_kod(self):
        kodlar = {"120", "120.01", "120.01.001", "320", "32001"}
        self.assertEqual(main.ust_kod("120.01.001", kodlar), "120.01")
        self.assertEqual(main.ust_kod("120.01", kodlar), "120")
        self.assertEqual(main.ust_kod("32001", kodlar), "320")
        self.assertIsNone(main.ust_kod("120", kodlar))

    def test_para(self):
        self.assertEqual(main.para("1.250.000,50"), D("1250000.50"))
        self.assertEqual(main.para("(1.000,00)"), D("-1000.00"))
        self.assertEqual(main.para("1.250.000"), D("1250000"))


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "mizan_eylul.csv", Path(tmp) / "a.xlsx", ORNEK / "mizan_agustos.csv")
        cls.b = {(x["kontrol"], x["kod"]): x for x in cls.s["bulgular"]}

    def test_denklik(self):
        self.assertEqual(self.s["denklik"]["net"], 0)
        hb, ha = self.s["denklik"]["hareket"]
        self.assertEqual(hb, ha)
        self.assertFalse(any(x["kontrol"] == "Denklik" for x in self.s["bulgular"]))

    def test_bicim_hatalari(self):
        self.assertIn(("Alt hesap toplamı", "120.01"), self.b)
        self.assertIn(("Satır aritmetiği", "335.01"), self.b)

    def test_ters_bakiyeler(self):
        self.assertEqual(self.b[("Ters bakiye", "360")]["seviye"], "Orta")
        for k in ("102.03", "120.01.014", "153.02", "320.01.007"):
            self.assertIn(("Alt hesap ters bakiye", k), self.b)
        self.assertNotIn(("Yön değişimi", "360"), self.b)            # ters bakiyede zaten var

    def test_mantik_ve_degisim(self):
        for anahtar in [("Kasa şişkinliği", "100"), ("Ortaklardan alacaklar", "131"), ("KDV mahsubu", "191/391"),
                        ("Yansıtma", "770/771"), ("Kümülatif azalma", "632"), ("Olağandışı değişim", "300"),
                        ("Yeni bakiye", "242")]:
            self.assertIn(anahtar, self.b)
        self.assertEqual(self.b[("Yansıtma", "770/771")]["tutar"], D("60000"))
        self.assertNotIn(("Yansıtma", "760/761"), self.b)
        self.assertNotIn(("Yeni bakiye", "191"), self.b)

    def test_denk_olmayan_ve_yil_basi(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "m.csv").write_text("Hesap Kodu;Hesap Adı;Bakiye\n100;Kasa;-50\n102;Banka;1000\n600;Satış;-900\n", encoding="utf-8")
            (t / "o.csv").write_text("Hesap Kodu;Hesap Adı;Bakiye\n102;Banka;9000\n600;Satış;-9000\n", encoding="utf-8")
            s = main.calistir(t / "m.csv", t / "a.xlsx", t / "o.csv")
        b = {(x["kontrol"], x["kod"]): x for x in s["bulgular"]}
        self.assertIn(("Denklik", "-"), b)
        self.assertEqual(b[("Ters bakiye", "100")]["seviye"], "Yüksek")
        self.assertNotIn(("Kümülatif azalma", "600"), b)               # yıl başı: 600 yarıdan fazla azaldı
        self.assertTrue(s["notlar"])


if __name__ == "__main__":
    unittest.main()
