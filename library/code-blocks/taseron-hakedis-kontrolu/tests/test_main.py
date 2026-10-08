import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class OrnekTest(unittest.TestCase):
    def setUp(self):
        with tempfile.TemporaryDirectory() as t:
            self.s = main.calistir(ORNEK / "sozlesme_birim_fiyatlar.csv", ORNEK / "taseron_hakedis_5.csv", Path(t) / "r.xlsx",
                                   ORNEK / "saha_metraji.csv", ORNEK / "onayli_hakedis_4.csv")
        self.b = {(x["poz"], x["kontrol"]) for x in self.s["bulgular"]}
        self.p = {main.poz_anahtari(x["poz"]): x for x in self.s["satirlar"]}

    def test_bulgular(self):
        for beklenen in [("15.180.1001", "Aritmetik"), ("15.180.1001", "Önceki miktar"), ("15 150 1005", "Saha metrajı aşımı"),
                         ("15.160.1003", "Birim fiyat"), ("EK-01", "Sözleşmede yok"), ("15.185.1005", "Sözleşme miktarı aşımı"),
                         ("Y.18.110", "Saha metrajı yok"), ("15.275.1101", "Hakedişte yok")]:
            self.assertIn(beklenen, self.b)
        self.assertNotIn(("15 150 1005", "Sözleşmede yok"), self.b)          # yazım farkına rağmen eşleşti

    def test_onay_tutarlari(self):
        self.assertEqual(self.p["151501005"]["onay_tutar"], Decimal("406250.00"))   # (945 − 820) × 3.250
        self.assertEqual(self.p["151601003"]["onay_tutar"], Decimal("413250.00"))   # 14,5 × 28.500 (sözleşme fiyatı)
        self.assertEqual(self.p["151801001"]["onay_bu"], Decimal("450"))            # talep edilenden fazlası onaylanmaz
        self.assertEqual(self.p["ek01"]["onay_tutar"], Decimal("0.00"))
        self.assertEqual(self.p["y18110"]["onay_tutar"], Decimal("0.00"))
        self.assertEqual(self.s["ozet"]["onay"], Decimal("1222700.00"))
        self.assertEqual(self.s["ozet"]["talep"], Decimal("1590350.00"))


class ElleTest(unittest.TestCase):
    def test_tolerans_ve_sozlesme_siniri(self):
        with tempfile.TemporaryDirectory() as t:
            t = Path(t)
            (t / "s.csv").write_text("Poz No;Birim;Birim Fiyat;Sözleşme Miktarı\nA;m2;10;100\n", encoding="utf-8")
            (t / "h.csv").write_text("Poz No;Birim;Birim Fiyat;Önceki Miktar;Bu Dönem Miktar;Toplam Miktar;Toplam Tutar\n"
                                     "A;m2;10;0;105;105;1050\n", encoding="utf-8")
            (t / "m.csv").write_text("Poz No;Kümülatif Miktar\nA;104\n", encoding="utf-8")
            s = main.calistir(t / "s.csv", t / "h.csv", t / "r.xlsx", t / "m.csv", tolerans=1)
            self.assertNotIn("Saha metrajı aşımı", {b["kontrol"] for b in s["bulgular"]})   # 105 ≤ 104 × 1,01
            self.assertEqual(s["satirlar"][0]["onay_toplam"], Decimal("104"))
            s = main.calistir(t / "s.csv", t / "h.csv", t / "r.xlsx", t / "m.csv", sozlesme_siniri=True)
            self.assertEqual(s["satirlar"][0]["onay_tutar"], Decimal("1000.00"))
            self.assertIn("Saha metrajı aşımı", {b["kontrol"] for b in s["bulgular"]})


if __name__ == "__main__":
    unittest.main()
