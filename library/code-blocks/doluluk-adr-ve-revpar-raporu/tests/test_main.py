import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
REZ = """Rezervasyon No;Giriş Tarihi;Çıkış Tarihi;Oda Adedi;Kanal;Gecelik Fiyat;Ücret Tipi;Durum
A;07.09.2026;09.09.2026;1;Online;1000;;Konakladı
B;07.09.2026;08.09.2026;1;Direkt;1500;;Konakladı
C;08.09.2026;09.09.2026;1;Direkt;0;Ücretsiz;Konakladı
D;07.09.2026;08.09.2026;1;Online;2000;;İptal
E;08.09.2025;09.09.2025;1;Online;900;;Konakladı
"""


class RezervasyonTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "r.csv").write_text(REZ, encoding="utf-8")
            cls.s = main.calistir(t / "r.csv", t / "o.xlsx", oda=2, baslangic=date(2026, 9, 7), bitis=date(2026, 9, 8))

    def test_gunluk(self):
        g1, g2 = self.s["gunluk"][date(2026, 9, 7)], self.s["gunluk"][date(2026, 9, 8)]
        self.assertEqual((g1["doluluk"], g1["adr"], g1["revpar"]), (1.0, 1250.0, 1250.0))     # A + B, iptal D hariç
        self.assertEqual((g2["doluluk"], g2["toplam_doluluk"], g2["adr"], g2["revpar"]), (0.5, 1.0, 1000.0, 500.0))  # C ücretsiz

    def test_donem(self):
        t = self.s["toplam"]
        self.assertEqual((t["doluluk"], t["revpar"]), (0.75, 875.0))
        self.assertAlmostEqual(t["adr"], 3500 / 3)
        self.assertAlmostEqual(t["revpar"], t["doluluk"] * t["adr"])          # RevPAR = doluluk × ADR

    def test_gecen_yil_haftanin_ayni_gunu(self):
        ly = self.s["gunluk"][date(2026, 9, 7)]["ly"]                         # 364 gün önce: 08.09.2025, Pazartesi
        self.assertEqual((ly["doluluk"], ly["adr"]), (0.5, 900.0))
        self.assertIsNone(self.s["gunluk"][date(2026, 9, 8)]["ly"])
        self.assertIsNone(self.s["aylar"][(2026, 9)]["ly"])                   # GY ayı eksik: karşılaştırma yok

    def test_kanal_ve_iptal(self):
        self.assertEqual({k: float(v["gelir"]) for k, v in self.s["kanallar"].items()}, {"Online": 2000.0, "Direkt": 1500.0})
        self.assertTrue(any("1 iptal" in u for u in self.s["uyarilar"]))


class GunlukRaporTesti(unittest.TestCase):
    def test_ooo_ve_overbooking(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "g.csv").write_text("Tarih;Satılan Oda;Ücretsiz Oda;Kullanım Dışı;Oda Geliri\n01.09.2026;80;2;10;240000\n"
                                     "02.09.2026;101;0;0;300000\n", encoding="utf-8")
            a = main.calistir(t / "g.csv", t / "o.xlsx", oda=100)
            b = main.calistir(t / "g.csv", t / "o.xlsx", oda=100, ooo_dus=True)
        self.assertEqual(a["gunluk"][date(2026, 9, 1)]["doluluk"], 0.8)
        self.assertAlmostEqual(b["gunluk"][date(2026, 9, 1)]["doluluk"], 80 / 90)
        self.assertEqual(a["gunluk"][date(2026, 9, 1)]["adr"], 3000.0)
        self.assertTrue(any("aşıyor" in u for u in a["uyarilar"]))


class OrnekTesti(unittest.TestCase):
    def test_ornek_tutarlilik(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "rezervasyonlar.csv", Path(tmp) / "o.xlsx", 40, date(2026, 9, 1), date(2026, 9, 30))
        t = s["toplam"]
        self.assertEqual(len(s["gunluk"]), 30)
        self.assertEqual(float(t["kullanilabilir"]), 1200)
        self.assertAlmostEqual(t["revpar"], t["doluluk"] * t["adr"])
        self.assertIsNotNone(s["aylar"][(2026, 9)]["ly"])
        self.assertTrue(all(g["satilan"] + g["comp"] + g["house"] <= 40 for g in s["gunluk"].values()))


if __name__ == "__main__":
    unittest.main()
