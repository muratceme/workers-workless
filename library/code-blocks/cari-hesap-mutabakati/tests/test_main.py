import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


def kayit(taraf, gun, belge, aciklama, borc=0, alacak=0):
    return main.Kayit(taraf, 0, date(2026, 9, gun), belge, main.tur_bul("", aciklama, belge), aciklama, D(borc), D(alacak))


class YardimciTesti(unittest.TestCase):
    def test_belge_ayni(self):
        self.assertTrue(main.belge_ayni("KAM2026000000123", "123"))         # e-Fatura no ↔ kısa sıra no
        self.assertTrue(main.belge_ayni("000123", "123"))
        self.assertFalse(main.belge_ayni("KAM2026000000123", "XYZ2026000000123"))   # farklı seri
        self.assertFalse(main.belge_ayni("5123", "123"))
        self.assertEqual(main.belge_anahtari("kam-2026/000000123"), "KAM2026000000123")

    def test_tur(self):
        self.assertEqual(main.tur_bul("", "EFT ödeme", ""), "Ödeme / tahsilat")
        self.assertEqual(main.tur_bul("", "Defter düzeltme", ""), "Diğer")         # 'eft' kelime içinde aranmaz
        self.assertEqual(main.tur_bul("İade Faturası", "", ""), "İade")
        self.assertEqual(main.tur_bul("", "Ağustos kur farkı", ""), "Kur farkı")
        self.assertEqual(main.tur_bul("", "", "ABC2026000000001"), "Fatura")

    def test_tutar_farki_nedeni(self):
        self.assertIn("%20", main.tutar_farki_nedeni(D(-60000), D(-72000)))       # KDV hariç / dahil
        # 10.000 + %20 KDV = 12.000; 5/10 tevkifat = 1.000 → karşı taraf 11.000 kaydetmiş
        self.assertIn("5/10", main.tutar_farki_nedeni(D(12000), D(11000)))
        self.assertIn("yuvarlama", main.tutar_farki_nedeni(D("100.00"), D("100.40")))

    def test_yon(self):
        biz = [kayit("Biz", 1, "A1", "fatura", alacak=100), kayit("Biz", 2, "A2", "fatura", alacak=50)]
        ters = [kayit("Karşı", 1, "A1", "fatura", borc=100), kayit("Karşı", 2, "A2", "fatura", borc=50)]
        duz = [kayit("Karşı", 1, "A1", "fatura", alacak=100)]
        self.assertTrue(main.yon_belirle(biz, ters))
        self.assertFalse(main.yon_belirle(biz, duz))


class UctanUcaTest(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "bizim_ekstre.csv", ORNEK / "karsi_taraf_ekstresi.csv", Path(tmp) / "m.xlsx",
                              bitis=date(2026, 9, 30))
        # Elle hesap: biz −134.000 (borçluyuz), karşı −217.850 → fark 83.850
        #   = iade 6.000 + yoldaki ödeme 40.000 + KDV farkı 12.000 + kur farkı 1.850 + yoldaki fatura 24.000
        self.assertTrue(s["ters"])
        self.assertEqual(s["bizim_bakiye"], D(-134000))
        self.assertEqual(s["karsi_bakiye"], D(-217850))
        self.assertEqual(s["fark"], D(83850))
        self.assertEqual(s["aciklanan"], s["fark"])
        self.assertFalse(s["mutabik"])
        turler = sorted((f[0], f[1].tur if f[1] else f[2].tur, f[3]) for f in s["farklar"])
        self.assertEqual(turler, sorted([
            ("Tutar farkı", "Fatura", D(12000)),
            ("Bizde var, karşıda yok", "İade", D(6000)),
            ("Bizde var, karşıda yok", "Ödeme / tahsilat", D(40000)),
            ("Karşıda var, bizde yok", "Kur farkı", D(1850)),
            ("Karşıda var, bizde yok", "Fatura", D(24000)),
        ]))
        yolda = [f for f in s["farklar"] if "yolda" in f[4]]
        self.assertEqual(len(yolda), 2)                     # 28.09 fatura ve 30.09 ödeme
        self.assertEqual(sum(1 for x in s["biz"] if x.eslesme == "Mutabık"), 6)

    def test_tarih_farki_ve_mukerrer(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "b.csv").write_text("Tarih;Belge No;Açıklama;Borç;Alacak\n01.09.2026;F1;fatura;1000;\n"
                                     "01.09.2026;F1;fatura;1000;\n05.09.2026;;EFT tahsilat;;500\n", encoding="utf-8")
            (t / "k.csv").write_text("Tarih;Belge No;Açıklama;Borç;Alacak\n01.09.2026;F1;fatura;;1000\n"
                                     "25.09.2026;;EFT ödeme;500;\n", encoding="utf-8")
            s = main.calistir(t / "b.csv", t / "k.csv", t / "m.xlsx")
        self.assertTrue(any("mükerrer" in u for u in s["uyarilar"]))
        self.assertIn(("Tarih farkı", D(0)), [(f[0], f[3]) for f in s["farklar"]])
        self.assertEqual(s["fark"], D(1000))                # mükerrer fatura bizde fazla
        self.assertEqual(s["aciklanan"], s["fark"])

    def test_mutabik(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "bizim_ekstre.csv", ORNEK / "bizim_ekstre.csv", Path(tmp) / "m.xlsx")
        # Aynı dosya iki kez: yön "aynı" algılanmalı ve her şey mutabık olmalı
        self.assertFalse(s["ters"])
        self.assertTrue(s["mutabik"])


if __name__ == "__main__":
    unittest.main()
