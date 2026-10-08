import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402
import mizan_cekirdek as mc  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


def grup(bolumler, bolum, harf):
    for b, gruplar, _ in bolumler:
        if b.startswith(bolum):
            return next(g for g in gruplar if g[0] == harf)


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "mizan_2026.csv", Path(tmp) / "a.xlsx", ORNEK / "mizan_2025.csv", "Test")
        cls.c = cls.s["Cari Dönem"]["t"]

    def test_denklik_ve_kar(self):
        self.assertEqual(self.c["aktif_top"], self.c["pasif_top"])
        self.assertEqual(self.c["aktif_top"], D("16300000"))
        self.assertEqual(self.c["donem_net"], D("2925000"))
        o = self.s["Önceki Dönem"]["t"]
        self.assertEqual((o["aktif_top"], o["pasif_top"], o["donem_net"]), (D("12700000"), D("12700000"), D("1800000")))
        self.assertEqual(self.s["Cari Dönem"]["uyarilar"], [])

    def test_virmanlar(self):
        v = {x["kod"]: x["hedef"] for x in self.s["Cari Dönem"]["virmanlar"]}
        self.assertEqual(v, {"102.03": "300", "120.03": "340"})
        self.assertEqual(grup(self.c["aktif"], "I.", "A")[3], D("3900000"))      # 50.000 kasa + 3.850.000 banka
        self.assertEqual(grup(self.c["pasif"], "III.", "A")[3], D("2650000"))    # 2.500.000 + 150.000 KMH
        self.assertEqual(grup(self.c["pasif"], "III.", "D")[3], D("50000"))

    def test_duzenleyici_ve_ad(self):
        dmv = grup(self.c["aktif"], "II.", "D")
        self.assertEqual(dmv[3], D("5000000"))                                  # 6.000.000 + 1.000.000 − 2.000.000
        self.assertIn(("257", "Birikmiş Amortismanlar (-)", D("-2000000")), dmv[2])
        self.assertEqual(grup(self.c["pasif"], "III.", "B")[2][0][:2], ("320", "Satıcılar"))   # yalnız alt hesaplar var

    def test_gelir_tablosu(self):
        g = {ad: t for _, ad, _, t, _ in self.c["gelir"]}
        self.assertEqual(g["NET SATIŞLAR"], D("23200000"))
        self.assertEqual(g["BRÜT SATIŞ KÂRI VEYA ZARARI"], D("7200000"))
        self.assertEqual(g["DÖNEM KÂRI VEYA ZARARI"], D("3900000"))

    def test_virmansiz_ve_kapanis_sonrasi(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "mizan_2026.csv", Path(tmp) / "a.xlsx", virman=False)
            self.assertEqual(s["Cari Dönem"]["t"]["aktif_top"], D("16100000"))
            t = Path(tmp) / "k.csv"
            t.write_text("Hesap Kodu;Hesap Adı;Bakiye\n100;Kasa;1000\n500;Sermaye;-800\n590;Dönem Net Kârı;-200\n", encoding="utf-8")
            k = main.calistir(t, Path(tmp) / "b.xlsx")["Cari Dönem"]
        self.assertEqual(k["t"]["donem_net"], D("200"))
        self.assertTrue(any("kapalı" in u for u in k["uyarilar"]))
        self.assertEqual(k["t"]["pasif_top"], D("1000"))


if __name__ == "__main__":
    unittest.main()
