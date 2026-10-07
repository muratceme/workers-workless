import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri" / "nakit_kalemleri.csv"


class AkisTesti(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmp.cleanup()

    def calistir(self, **kw):
        return main.calistir(ORNEK, D(850000), date(2026, 10, 5), Path(self.tmp.name) / "n.xlsx", **kw)

    def test_bakiye_kimligi(self):
        s = self.calistir()
        self.assertEqual(s["kapanis"][-1], D(850000) + sum(s["giris"], D(0)) - sum(s["cikis"], D(0)))

    def test_tekrarlayan_kalemlerin_gecmisi_alinmaz(self):
        s = self.calistir()
        maaslar = [a for a in s["akislar"] if a.kategori == "Maaş"]
        # 05.10-03.01 arası: 05.10, 05.11, 05.12 maaşları (Ocak'tan bu yana geçmiş aylar dahil edilmez)
        self.assertEqual([a.tarih for a in maaslar], [date(2026, 10, 5), date(2026, 11, 5), date(2026, 12, 5)])
        kredi = [a.tarih for a in s["akislar"] if a.kategori == "Kredi"]
        self.assertEqual(kredi, [date(2026, 10, 20), date(2026, 11, 20), date(2026, 12, 20)])   # bitiş 20.12

    def test_gecikmis_kalemler(self):
        s = self.calistir(gecikmis_oran=D("0.5"))
        ilk_hafta = [a for a in s["akislar"] if a.tarih == date(2026, 10, 5)]
        gecmis_tahsilat = sum((a.tutar for a in ilk_hafta if a.tur == "Giriş" and a.not_), D(0))
        self.assertEqual(gecmis_tahsilat, (D(20000) + D(64500)) / 2)
        self.assertTrue(any(a.tur == "Çıkış" and a.tutar == D(42000) for a in ilk_hafta))

    def test_tahsilat_gecikmesi(self):
        s = self.calistir(gecikme=10)
        dl = next(a for a in s["akislar"] if "ORN099" in a.aciklama)
        self.assertEqual(dl.tarih, date(2026, 10, 20))
        cek = next(a for a in s["akislar"] if "Ç-48211" in a.aciklama)
        self.assertEqual(cek.tarih, date(2026, 10, 23))               # çek: müşteri gecikmesi uygulanmaz

    def test_minimum_esik(self):
        s = self.calistir(minimum=D(250000))
        for i in s["acik_haftalar"]:
            self.assertLess(s["kapanis"][i], D(250000))


if __name__ == "__main__":
    unittest.main()
