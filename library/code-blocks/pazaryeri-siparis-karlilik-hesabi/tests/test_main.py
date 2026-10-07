import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


def ornek():
    with tempfile.TemporaryDirectory() as tmp:
        return main.calistir(ORNEK / "siparisler.csv", ORNEK / "urun_maliyetleri.csv", ORNEK / "kanallar.json", Path(tmp) / "p.xlsx")


def satir(s, siparis, kod):
    return next(x for x in s["satirlar"] if x.siparis == siparis and x.kod == kod).sonuc


class HesapTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.s = ornek()

    def test_komisyon_kdv_dahil(self):
        # 1002: 149,90 TL (KDV %20 dahil), komisyon %18 KDV dahil tutardan = 26,982 → KDV hariç 22,485
        # kargo 2 desi = 66 → 55; hizmet 10,19 → 8,4917; maliyet 45
        # kâr = 124,9167 − 22,485 − 55 − 8,4917 − 45 = −6,06
        r = satir(self.s, "1002", "KP-010")
        self.assertEqual(main.yuvarla(r["komisyon"]), D("22.49"))
        self.assertEqual(main.yuvarla(r["kar"]), D("-6.06"))
        self.assertEqual(main.yuvarla(r["stopaj"]), D("1.25"))          # %1 × KDV hariç satış

    def test_komisyon_kdv_haric(self):
        # 2001: 949 TL, komisyon %15 × 790,83 = 118,625 (+KDV), kargo 60 → 50, maliyet 380 → kâr 242,21
        r = satir(self.s, "2001", "KZ-300")
        self.assertEqual(main.yuvarla(r["komisyon"]), D("118.63"))
        self.assertEqual(main.yuvarla(r["kar"]), D("242.21"))
        self.assertEqual(main.yuvarla(r["odeme"]), D("738.74"))         # 949 − 142,35 − 60 − 7,91

    def test_iade_ve_iptal(self):
        r = satir(self.s, "1005", "TS-001")
        self.assertEqual(main.yuvarla(r["kar"]), D("-98.49"))          # (54 + 54) / 1,2 + 8,49 kayıp
        self.assertEqual(satir(self.s, "1006", "KP-010")["kar"], D(0))

    def test_cok_kalemli_siparis_kargo_dagitimi(self):
        a, b = satir(self.s, "1007", "TS-001"), satir(self.s, "1007", "KP-010")
        self.assertEqual(main.yuvarla(a["kargo"] + b["kargo"]), D("55.00"))   # 1+2 desi = 66 TL tek kargo
        self.assertGreater(a["kargo"], b["kargo"])                          # satış tutarı oranında

    def test_satici_indirimi(self):
        r = satir(self.s, "1004", "KZ-300")
        self.assertEqual(main.yuvarla(r["net_satis"]), D("707.50"))     # (899 − 50) / 1,2

    def test_basabas_ve_uyarilar(self):
        kp = self.s["urunler"][("Trendyol", "KP-010")]
        self.assertEqual(main.yuvarla(kp["basabas"]), D("158.77"))     # (76,19/1,2 + 45) / (1/1,2 − 0,18/1,2)
        self.assertTrue(any("XX-999" in u for u in self.s["uyarilar"]))

    def test_oran_okuma(self):
        for x in ("%20", "20", "0,20", "0.20"):
            self.assertEqual(main.oran(x), D("0.20"))


if __name__ == "__main__":
    unittest.main()
