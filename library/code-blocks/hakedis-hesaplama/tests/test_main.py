import json
import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


def hesapla(ayar_degisikligi=None, n=3):
    a = main.ayarlari_oku(ORNEK / "sozlesme.json")
    a.update(ayar_degisikligi or {})
    pozlar = main.pozlari_oku(ORNEK / "pozlar.csv")
    metraj, _ = main.metraj_oku(ORNEK / "metraj.csv", pozlar)
    return main.icmal(pozlar, metraj, a, n)


class IcmalTesti(unittest.TestCase):
    def test_ornek_3_nolu_hakedis(self):
        # Elle hesap:
        #   önceki iş: 4000×180 + 400×2600 + 50×6000 + 2300×100           = 2.290.000
        #   bu iş   : 120×180 + 600,5×2600 + 70,25×6000 + 1700×100        = 2.174.400
        #   + fiyat farkı 85.000 → bu hakediş                             = 2.259.400
        #   KDV %20 = 451.880 · tevkifat 4/10 = 180.752 (KDV dahil bedel 7,44 mn ≥ 5 mn)
        #   stopaj %5 = 112.970 · teminat %5 = 112.970 · avans mahsubu %10 = 225.940 · elektrik 12.500
        ic = hesapla()
        self.assertEqual(ic["onceki_is"], D("2290000.00"))
        self.assertEqual(ic["bu_is"], D("2174400.00"))
        self.assertEqual(ic["bu_hakedis"], D("2259400.00"))
        self.assertEqual(ic["kdv"], D("451880.00"))
        self.assertEqual(ic["tevkifat"], D("180752.00"))
        self.assertEqual(dict(ic["kesintiler"])["Gelir / kurumlar vergisi stopajı (%5)"], D("112970.00"))
        self.assertEqual(ic["avans_mahsup"], D("225940.00"))
        self.assertEqual(ic["kesinti_toplami"], D("464380.00"))
        self.assertEqual(ic["odenecek"], D("2066148.00"))

    def test_tevkifat_esigi(self):
        # KDV dahil bedel 5 mn altında ve belirlenmiş alıcı değilse tevkifat yok
        self.assertEqual(hesapla({"sozlesme_bedeli": "4000000"})["tevkifat"], D(0))
        # Belirlenmiş alıcıya her tutarda 4/10
        self.assertEqual(hesapla({"sozlesme_bedeli": "4000000", "belirlenmis_alici": True})["tevkifat"], D("180752.00"))
        # Elle oran ve "yok"
        self.assertEqual(hesapla({"kdv_tevkifati": "yok"})["tevkifat"], D(0))
        self.assertEqual(hesapla({"kdv_tevkifati": "2/10"})["tevkifat"], D("90376.00"))

    def test_stopaj(self):
        ic = hesapla({"yillara_yaygin": False})
        self.assertFalse(any("stopaj" in k[0] for k in ic["kesintiler"]))
        ic = hesapla({"stopaj_orani": "0.01"})                     # demiryolu / gemi / nükleer
        self.assertEqual(dict(ic["kesintiler"])["Gelir / kurumlar vergisi stopajı (%1)"], D("22594.00"))

    def test_avans_tavani(self):
        # Toplam mahsup hiçbir zaman verilen avansı aşmamalı
        toplam = sum(hesapla({"avans_tutari": "300000"}, n)["avans_mahsup"] for n in (1, 2, 3))
        self.assertLessEqual(toplam, D(300000))
        ic = hesapla({"avans_tutari": "100000"}, 3)
        self.assertEqual(ic["avans_mahsup"], D("36441.94"))         # 2.259.400 × 100.000 / 6.200.000

    def test_ilk_hakedis_ve_uyarilar(self):
        ic = hesapla(n=1)
        self.assertEqual((ic["onceki_is"], ic["bu_hakedis"]), (D(0), D("530000.00")))
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "sozlesme.json", ORNEK / "pozlar.csv", ORNEK / "metraj.csv", None, Path(tmp) / "h.xlsx")
        self.assertEqual(s["icmal"]["n"], 3)
        self.assertTrue(any("P-99" in u for u in s["uyarilar"]))
        self.assertTrue(any("P-01" in u and "aşıyor" in u for u in s["uyarilar"]))

    def test_sozlesme_json_gecerli(self):
        a = json.loads((ORNEK / "sozlesme.json").read_text(encoding="utf-8"))
        self.assertTrue(set(a) <= set(main.VARSAYILAN))


if __name__ == "__main__":
    unittest.main()
