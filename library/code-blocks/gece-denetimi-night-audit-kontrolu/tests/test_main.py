import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

O = KLASOR / "ornek_veri"


def ornek(t, **k):
    return main.calistir(O / "folyolar.csv", Path(t) / "r.xlsx", date(2026, 10, 7), O / "fiyatlar.csv", O / "oda_durumu.csv",
                         O / "tahsilat.csv", O / "kasa_sayim.csv", **k)


class OrnekTest(unittest.TestCase):
    def test_bulgular(self):
        with tempfile.TemporaryDirectory() as t:
            s = ornek(t, bakiye_limiti=15000, rapor_oda_geliri=25350)
        k = {(b["ref"], b["kontrol"]) for b in s["bulgular"]}
        for beklenen in [("109 / F0990", "Açık folyo"), ("107 / F1007", "Çıkış tarihi geçmiş"), ("105 / F1005", "Oda ücreti postalanmamış"),
                         ("102 / F1002", "Fiyat farkı"), ("105", "Skip"), ("109", "Sleep"), ("103 / F1003", "Bugün çıkışı var"),
                         ("104 / F1004", "Bakiye limiti"), ("106 / F1006", "Ücretsiz oda"), ("F9999", "Folyosuz tahsilat"),
                         ("Murat / Nakit", "Kasa farkı"), ("Oda geliri", "Gelir raporu farkı")]:
            self.assertIn(beklenen, k)
        self.assertNotIn(("101 / F1001", "Fiyat farkı"), k)
        say = {sv: sum(1 for b in s["bulgular"] if b["seviye"] == sv) for sv in main.SEVIYE_SIRA}
        self.assertEqual(say, {"Hata": 5, "Yüksek": 3, "Dikkat": 5, "Bilgi": 1})

    def test_ozet(self):
        with tempfile.TemporaryDirectory() as t:
            s = ornek(t)
        o = s["ozet"]
        self.assertEqual(o["oda_geliri"], Decimal("21850"))       # 3500+4500+3000+4200+0+3500+3150
        self.assertEqual(o["dolu_oda"], 8)
        self.assertEqual(o["toplam_oda"], 9)                      # 110 OOO
        self.assertEqual(o["adr"], Decimal("3641.67"))            # 6 ücretli oda
        self.assertEqual(o["revpar"], Decimal("2427.78"))
        self.assertEqual(o["kasa_farki"], Decimal("-150"))
        self.assertFalse(any(b["kontrol"] in ("Bakiye limiti", "Gelir raporu farkı") for b in s["bulgular"]))

    def test_kasa_tolerans(self):
        with tempfile.TemporaryDirectory() as t:
            s = ornek(t, kasa_tolerans=200)
        self.assertFalse(any(b["kontrol"] == "Kasa farkı" for b in s["bulgular"]))


class DurumTest(unittest.TestCase):
    def test_durum(self):
        self.assertEqual(main.durum_sinifi("In-House"), "konakliyor")
        self.assertEqual(main.durum_sinifi("Checked Out"), "cikis")
        self.assertEqual(main.durum_sinifi("No Show"), "noshow")
        self.assertEqual(main.durum_sinifi("ÇIKIŞ YAPTI"), "cikis")


if __name__ == "__main__":
    unittest.main()
