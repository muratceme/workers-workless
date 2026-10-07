import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class YardimciTesti(unittest.TestCase):
    def test_gelir_gider(self):
        self.assertTrue(main.gelir_mi("600.01"))
        self.assertTrue(main.gelir_mi("649"))
        self.assertFalse(main.gelir_mi("610"))          # satış indirimleri
        self.assertFalse(main.gelir_mi("770.01"))
        self.assertTrue(main.gelir_mi("999", "Gelir"))  # Tür sütunu önceliklidir

    def test_onek(self):
        adaylar = ["770", "770.02", "770.0"]
        self.assertEqual(main.onek_bul("770.02.001", adaylar), "770.02")
        self.assertEqual(main.onek_bul("770.05", adaylar), "770")
        self.assertIsNone(main.onek_bul("7700", ["770"]))   # hane sınırı

    def test_ay(self):
        self.assertEqual(main.ay_coz("Eylül"), 9)
        self.assertEqual(main.ay_coz("2026-09"), 9)
        self.assertEqual(main.ay_coz("09.2026"), 9)
        self.assertEqual(main.ay_coz("30.09.2026"), 9)

    def test_yon_ve_onem(self):
        self.assertEqual(main.sapma_yorumu(True, D(5)), "Lehte")
        self.assertEqual(main.sapma_yorumu(False, D(5)), "Aleyhte")
        self.assertFalse(main.onemli(D(9000), D(10000), D(10), D(10000)))     # tutar eşiği altında
        self.assertFalse(main.onemli(D(20000), D(1000000), D(10), D(10000)))  # yüzde eşiği altında
        self.assertTrue(main.onemli(D(12000), D(0), D(10), D(10000)))          # bütçesiz


class UctanUcaTest(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "butce_2026.csv", ORNEK / "gerceklesen_muavin.csv", Path(tmp) / "r.xlsx")
        self.assertEqual((s["yil"], s["ay"]), (2026, 9))                     # 2025 kaydı elenir, son ay Eylül
        self.assertTrue(s["uyarilar"])
        x = {(r["hesap"], r["mm"]): r for r in s["satirlar"]}
        satis = x[("600.01", "Satış")]
        self.assertEqual(satis["f_ay"], D(870000))                          # alacak − borç (iptal düşülür)
        self.assertEqual(satis["fark_ytd"], D(-90000))
        self.assertEqual(satis["yon_ay"], "Aleyhte")
        self.assertTrue(satis["onemli_ay"] and not satis["onemli_ytd"])
        self.assertIn("zamanlama", satis["not"])
        kira = x[("770.02", "Genel Yönetim")]
        self.assertEqual(kira["f_ytd"], D(720000))                          # 770.02.001 → 770.02
        dan = x[("770.03", "Genel Yönetim")]
        self.assertEqual(dan["fark_ytd"], D(70000))
        self.assertEqual(dan["tahmin"], D(310000))                          # 250.000 + 3 × 20.000
        reklam = x[("760.02", "Pazarlama")]
        self.assertEqual((reklam["b_ytd"], reklam["f_ytd"]), (D(550000), D(555000)))
        self.assertFalse(reklam["onemli_ytd"])
        faiz = x[("780.01", "Finansman")]
        self.assertFalse(faiz["onemli_ay"])                                  # −8.000 < 10.000 TL eşiği
        self.assertEqual(faiz["yon_ay"], "Lehte")
        temsil = x[("770.09.001", "Genel Yönetim")]
        self.assertIn("Bütçelenmemiş", temsil["not"])
        self.assertEqual({(r["hesap"]) for r in s["onemli"]}, {"600.01", "770.03", "770.09.001", "760.02"})

    def test_uzun_bicim_ve_ay_secimi(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "b.csv").write_text("Hesap Kodu;Ay;Tutar\n632;1;100\n632;2;100\n", encoding="utf-8")
            (t / "g.csv").write_text("Hesap Kodu;Ay;Tutar\n632;Ocak;150\n632;Şubat;80\n", encoding="utf-8")
            s = main.calistir(t / "b.csv", t / "g.csv", t / "r.xlsx", ay=1, esik_tutar=D(10))
        r = s["satirlar"][0]
        self.assertEqual((r["fark_ay"], r["fark_ytd"], r["tahmin"]), (D(50), D(50), D(250)))   # tahmin = Ocak fiili 150 + Şubat bütçesi 100
        self.assertTrue(r["onemli_ay"])


if __name__ == "__main__":
    unittest.main()
