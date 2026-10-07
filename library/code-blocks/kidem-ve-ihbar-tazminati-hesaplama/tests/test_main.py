import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402


class HizmetSuresiTesti(unittest.TestCase):
    def test_gun_dahil(self):
        self.assertEqual(main.hizmet_suresi(date(2019, 3, 15), date(2026, 9, 30)), (7, 6, 16))
        self.assertEqual(main.hizmet_suresi(date(2020, 1, 1), date(2020, 12, 31)), (1, 0, 0))
        # 31 Ocak + 1 ay = 28 Şubat; 28 Şubat → 2 Mart (çıkış günü dahil) = 2 gün; toplam 30 gün
        self.assertEqual(main.hizmet_suresi(date(2026, 1, 31), date(2026, 3, 1)), (0, 1, 2))


class KidemTesti(unittest.TestCase):
    def test_bagimsiz_ornek(self):
        # Bağımsız örnek: 40.000 TL × (5 yıl + 3/12 + 12/365) = 211.315,07
        # (kaynakta gün kısmı 1.315,08 yazılmış; 40.000 × 12 / 365 = 1.315,068 → 1.315,07)
        s = main.hesapla("X", "2020-01-01", "2025-04-12", 40000)
        self.assertEqual(s.hizmet, "5 yıl 3 ay 12 gün")
        self.assertEqual(s.kidem_brut, D("211315.07"))
        self.assertEqual(s.kidem_damga, D("1603.88"))     # 211.315,07 × binde 7,59

    def test_tavan_doneme_gore(self):
        s1 = main.hesapla("X", "2016-01-01", "2026-06-30", 100000)
        s2 = main.hesapla("X", "2016-01-01", "2026-07-01", 100000)
        self.assertEqual(s1.kidem_tavani, D("64948.77"))
        self.assertEqual(s2.kidem_tavani, D("73729.87"))
        self.assertEqual(s1.kidem_esas_ucret, D("64948.77"))

    def test_bir_yildan_az(self):
        s = main.hesapla("X", "2026-01-01", "2026-11-30", 50000)
        self.assertEqual(s.kidem_brut, D("0.00"))
        self.assertIn("1 yıldan az", s.notlar)

    def test_giydirilmis_ucret(self):
        s = main.hesapla("X", "2020-01-01", "2026-09-30", 60000, yan_odeme=4500, ikramiye_yillik=60000)
        self.assertEqual(s.giydirilmis_brut, D("69500.00"))


class IhbarTesti(unittest.TestCase):
    def test_sureler(self):
        cases = [("2026-01-01", "2026-05-31", 2), ("2026-01-01", "2026-07-15", 4),
                 ("2024-01-01", "2026-01-15", 6), ("2020-01-01", "2026-01-15", 8)]
        for giris, cikis, hafta in cases:
            self.assertEqual(main.hesapla("X", giris, cikis, 30000).ihbar_suresi_hafta, hafta, (giris, cikis))

    def test_ihbar_vergileri(self):
        s = main.hesapla("X", "2020-01-01", "2026-03-31", 30000, kumulatif_matrah=0)
        self.assertEqual(s.ihbar_brut, D("56000.00"))      # 30.000 / 30 × 56 gün
        self.assertEqual(s.ihbar_gelir_vergisi, D("8400.00"))  # %15 dilim
        self.assertEqual(s.ihbar_damga, D("425.04"))


class UctanUcaTest(unittest.TestCase):
    def test_ornek_liste(self):
        with tempfile.TemporaryDirectory() as tmp:
            kayitlar = main.liste_oku(KLASOR / "ornek_veri" / "cikislar.csv")
            sonuclar = [main.hesapla(**k) for k in kayitlar]
            main.rapor_yaz(sonuclar, Path(tmp) / "t.xlsx")
            by = {s.ad: s for s in sonuclar}
            self.assertEqual(by["Mehmet Kaya"].kidem_brut, D("0.00"))
            self.assertEqual(by["Zeynep Demir"].ihbar_brut, D("0.00"))
            self.assertEqual(by["Ali Veli"].toplam_net, D("0.00"))
            self.assertGreater(by["Ayşe Yılmaz"].kidem_net, 0)


if __name__ == "__main__":
    unittest.main()
