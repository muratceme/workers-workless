import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
KURLAR = {"USD": D("41.20"), "EUR": D("48.05")}


def ornek(**kw):
    with tempfile.TemporaryDirectory() as tmp:
        return main.calistir(ORNEK / "teklifler.csv", Path(tmp) / "t.xlsx", KURLAR, D("0.40"), **kw)


class HesapTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.s = ornek()

    def test_bugunku_deger(self):
        # Örnek Endüstriyel: 19.000 + 14.000 + 99.500 (98.000 + 1.500 nakliye) + 19.000 = 151.500
        # 60 gün vade, yıllık %40: 151.500 / (1 + 0,40 × 60/365) = 142.152,96
        d = self.s["tedarikciler"]["Örnek Endüstriyel Ltd. Şti."]
        self.assertEqual(main.yuvarla(d["tutar"]), D("151500.00"))
        self.assertEqual(main.yuvarla(d["bd_tutar"]), D("142152.96"))

    def test_doviz(self):
        # 2,10 × 41,20 × 200 + 3,20 × 41,20 × 100 + 215 × 41,20 × 10 + 4.500 + 0,80 × 41,20 × 500 = 140.048
        self.assertEqual(main.yuvarla(self.s["tedarikciler"]["Deneme Global Trading"]["bd_tutar"]), D("140048.00"))

    def test_kalem_en_iyi(self):
        en = {k: t.tedarikci for k, t in self.s["en_iyi"].items()}
        self.assertEqual(en, {"K-01": "Kurgu Rulman A.Ş.", "K-02": "Kurgu Rulman A.Ş.",
                              "K-03": "Deneme Global Trading", "K-04": "Hayali Hırdavat"})

    def test_teknik_uygunsuz_ve_kapsam(self):
        h = self.s["tedarikciler"]["Hayali Hırdavat"]
        self.assertEqual(h["eksik"], ["K-03"])                      # VG 68 yağ uygun değil
        self.assertEqual(h["kapsam"], 0.75)
        self.assertEqual(self.s["oneri_tek"], "Örnek Endüstriyel Ltd. Şti.")   # kapsamı tam olanlar arasında en yüksek puan

    def test_puan(self):
        # fiyat 130.617,5 / 142.152,96 → 91,89 · termin 3/7 · vade 60/90 · kalite 85/90 · garanti 12/24
        d = self.s["tedarikciler"]["Örnek Endüstriyel Ltd. Şti."]
        self.assertAlmostEqual(d["puanlar"]["termin"], 300 / 7, places=6)
        self.assertAlmostEqual(d["puanlar"]["vade"], 6000 / 90, places=6)
        self.assertAlmostEqual(d["toplam_puan"], 79.04, delta=0.01)

    def test_asiri_dusuk(self):
        self.assertTrue(any("Hayali Hırdavat · İş eldiveni" in u and "aşırı düşük" in u for u in self.s["uyarilar"]))

    def test_agirlik_degisimi(self):
        s = ornek(agirlik={"fiyat": 100, "termin": 0, "vade": 0, "kalite": 0, "garanti": 0})
        self.assertEqual(s["oneri_tek"], "Deneme Global Trading")     # yalnız fiyat: en düşük bugünkü değer

    def test_eksik_kur(self):
        with tempfile.TemporaryDirectory() as tmp, self.assertRaises(SystemExit):
            main.calistir(ORNEK / "teklifler.csv", Path(tmp) / "t.xlsx", {})


if __name__ == "__main__":
    unittest.main()
