import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class NormalTesti(unittest.TestCase):
    def test_turkce_ve_transliterasyon(self):
        k = main.ad_kelimeleri
        self.assertEqual(k("Hüseyin Şükrü Öztürk"), ["huseyin", "sukru", "ozturk"])
        self.assertEqual(k("HUSSEIN"), k("Hüseyin"))
        self.assertEqual(k("Mohammed"), k("Mehmet"))
        self.assertEqual(k("Dr. Ali Bin Yusuf"), ["ali", "yusuf"])
        self.assertEqual(k("Örnek Ticaret A.Ş.", sirket=True), ["ornek", "ticaret"])

    def test_benzerlik_sira_bagimsiz(self):
        b = main.benzerlik
        self.assertEqual(b(main.ad_kelimeleri("Yılmaz Ahmet"), main.ad_kelimeleri("Ahmet Yılmaz")), 100.0)
        self.assertLess(b(main.ad_kelimeleri("Ahmet Yılmaz"), main.ad_kelimeleri("Elif Kaya")), 50)
        self.assertGreater(b(main.ad_kelimeleri("M. Kurgusali"), main.ad_kelimeleri("Mohammed Kurgusali")), 85)

    def test_tarih(self):
        self.assertEqual(main.iso("12.04.1979"), main.iso("1979-04-12"))


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as t:
            cls.s = main.calistir(ORNEK / "musteriler.csv", [ORNEK / "ornek_liste.xml", ORNEK / "ornek_liste.csv"], Path(t) / "a.xlsx")
        cls.e = {x["musteri"]["no"]: x for x in cls.s["eslesmeler"]}

    def test_bm_xml_okuma(self):
        bm = [k for k in self.s["liste"] if k["liste"] == "BM Konsolide Liste"]
        self.assertEqual(len(bm), 3)
        self.assertIn("Abu Yousef Kurgusali", bm[0]["diger"])
        self.assertEqual(bm[0]["dogum"], ["1979-04-12"])
        self.assertEqual(bm[2]["tur"], "Kuruluş")

    def test_eslesmeler(self):
        self.assertEqual(self.e["C001"]["seviye"], "Güçlü eşleşme")
        self.assertIn("doğum tarihi eşleşiyor", self.e["C001"]["notlar"])
        self.assertEqual(self.e["C003"]["seviye"], "Güçlü eşleşme")                 # Ltd. Şti. ↔ LLC
        self.assertEqual(self.e["C002"]["seviye"], "Olası eşleşme")                 # ad aynı, doğum yılı farklı
        self.assertEqual(self.e["C007"]["seviye"], "Güçlü eşleşme")                 # "ve" ve şirket ekleri yok sayılır
        for no in ("C004", "C005", "C008"):
            self.assertNotIn(no, self.e)


if __name__ == "__main__":
    unittest.main()
