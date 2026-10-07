import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402


class KuralTesti(unittest.TestCase):
    """4857 s. Kanun md. 53."""

    def test_hizmet_yili_kademeleri(self):
        self.assertEqual([main.yillik_izin_gunu(n, 30) for n in (1, 5, 6, 14, 15, 20)], [14, 14, 20, 20, 26, 26])

    def test_yas_kurali(self):
        self.assertEqual(main.yillik_izin_gunu(1, 18), 20)
        self.assertEqual(main.yillik_izin_gunu(1, 19), 14)
        self.assertEqual(main.yillik_izin_gunu(2, 50), 20)
        self.assertEqual(main.yillik_izin_gunu(16, 55), 26)     # 26 zaten 20'den fazla

    def test_yer_alti(self):
        self.assertEqual(main.yillik_izin_gunu(1, 30, yer_alti=True), 18)


class HesapTesti(unittest.TestCase):
    def test_yas_hakedis_tarihinde_degerlendirilir(self):
        # 50 yaşına 2026'da giren, 3. yılında (2026) 20 gün hak eder; öncekiler 14
        c = main.hesapla(main.Calisan("X", date(2023, 3, 1), dogum=date(1976, 1, 10)), date(2026, 12, 31))
        self.assertEqual([h.gun for h in c.hakedisler], [14, 14, 20])

    def test_29_subat(self):
        self.assertEqual(main.yil_ekle(date(2024, 2, 29), 1), date(2025, 2, 28))

    def test_ozet_ve_izin_ucreti(self):
        c = main.hesapla(main.Calisan("X", date(2019, 3, 15), dogum=date(1990, 6, 12),
                                      kullanilan=D(30), brut=D(60000)), date(2026, 9, 30))
        o = main.ozet(c, date(2026, 9, 30))
        self.assertEqual(o["Tamamlanan Hizmet Yılı"], 7)
        self.assertEqual(o["Toplam Hakediş"], 5 * 14 + 2 * 20)          # 110
        self.assertEqual(o["Kalan"], 80)
        self.assertEqual(o["Kullanılmayan İzin Ücreti (Brüt)"], 160000.0)  # 60.000 / 30 × 80
        self.assertEqual(o["Sonraki Hakediş Tarihi"], date(2027, 3, 15))

    def test_bir_yil_dolmadan(self):
        c = main.hesapla(main.Calisan("X", date(2026, 2, 1)), date(2026, 9, 30))
        self.assertEqual(c.hakedisler, [])


class UctanUcaTest(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            ref = date(2026, 9, 30)
            calisanlar = [main.hesapla(c, ref) for c in main.liste_oku(KLASOR / "ornek_veri" / "personel.csv")]
            ozetler = {o["Ad Soyad"]: o for o in main.rapor_yaz(calisanlar, ref, Path(tmp) / "i.xlsx")}
            self.assertEqual(ozetler["Can Öztürk"]["Son Hakediş (Gün)"], 20)   # 2025'te 17 yaşında
            self.assertEqual(ozetler["Hasan Çelik"]["Son Hakediş (Gün)"], 30)  # 20. yıl 26 + yer altı 4
            self.assertEqual(ozetler["Mehmet Kaya"]["Toplam Hakediş"], 0)


if __name__ == "__main__":
    unittest.main()
