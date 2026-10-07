import sys
import tempfile
import unittest
from collections import OrderedDict
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class MizanTesti(unittest.TestCase):
    def test_hiyerarsi(self):
        m, uyarilar = main.mizan_oku(ORNEK / "mizan_2026.csv")
        self.assertEqual(m["102"], D("3850000"))       # ana hesap satırı; 102.01 + 102.02 tekrar eklenmez
        self.assertEqual(m["320"], D("-3000000"))      # ana satır yok: yapraklar (320.01.001, .002, 320.02) toplanır
        self.assertEqual(uyarilar, [])                 # mizan denk

    def test_kapanmis_mizan(self):
        with tempfile.TemporaryDirectory() as tmp:
            yol = Path(tmp) / "m.csv"
            yol.write_text("Hesap Kodu;Borç Bakiye;Alacak Bakiye\n102;1000;\n500;;600\n590;;400\n", encoding="utf-8")
            t, u = main.tablolar(main.mizan_oku(yol)[0])
        self.assertFalse(t["acik_6"])
        self.assertEqual(t["bilanco"]["Dönem net kârı / zararı"], D(400))
        self.assertEqual(t["bilanco"]["AKTİF TOPLAMI"], t["bilanco"]["PASİF TOPLAMI"])
        self.assertTrue(any("kapalı" in x for x in u))


class RasyoTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(OrderedDict([("2025", ORNEK / "mizan_2025.csv"), ("2026", ORNEK / "mizan_2026.csv")]),
                                  Path(tmp) / "r.xlsx")
        cls.t = cls.s["donemler"]["2026"]
        cls.r = cls.t["rasyolar"]

    def test_tablolar(self):
        gt, bl = self.t["gelir"], self.t["bilanco"]
        self.assertEqual(gt["Net satışlar"], D(23200000))             # 24 mn − 0,8 mn indirim/iade
        self.assertEqual(gt["Brüt satış kârı"], D(7200000))
        self.assertEqual(gt["Esas faaliyet kârı"], D(4500000))
        self.assertEqual(gt["Dönem kârı (vergi öncesi)"], D(3900000))
        self.assertEqual(gt["Dönem net kârı"], D(2925000))
        self.assertEqual(bl["AKTİF TOPLAMI"], D(16100000))
        self.assertEqual(bl["PASİF TOPLAMI"], D(16100000))
        self.assertEqual(bl["KISA VADELİ YABANCI KAYNAKLAR"], D(6675000))   # 371 peşin vergi düşülür
        self.assertEqual(self.s["uyarilar"], [])

    def test_likidite_ve_yapi(self):
        self.assertAlmostEqual(self.r["Cari oran"], 11100000 / 6675000)
        self.assertAlmostEqual(self.r["Asit-test oranı"], 8600000 / 6675000)
        self.assertAlmostEqual(self.r["Nakit oranı"], 3900000 / 6675000)
        self.assertAlmostEqual(self.r["Kaldıraç oranı"], 8675000 / 16100000)

    def test_faaliyet_ortalama_bakiye(self):
        self.assertAlmostEqual(self.r["Ticari alacak devir hızı"], 23200000 / 3950000)    # (4,5 + 3,4) / 2
        self.assertAlmostEqual(self.r["Stok devir hızı"], 16000000 / 2250000)            # (2,5 + 2,0) / 2
        self.assertAlmostEqual(self.r["Ticari borç devir hızı"], 16000000 / 2750000)     # (3,0 + 2,5) / 2
        nds = 365 * 3950000 / 23200000 + 365 * 2250000 / 16000000 - 365 * 2750000 / 16000000
        self.assertAlmostEqual(self.r["Nakit dönüşüm süresi"], nds)

    def test_karlilik(self):
        self.assertAlmostEqual(self.r["Net kâr marjı"], 2925000 / 23200000)
        self.assertAlmostEqual(self.r["Özkaynak kârlılığı (ROE)"], 2925000 / ((7425000 + 4800000) / 2))
        self.assertAlmostEqual(self.r["Aktif kârlılığı (ROA)"], 2925000 / ((16100000 + 12700000) / 2))
        self.assertAlmostEqual(self.r["Faiz karşılama oranı"], 4700000 / 800000)

    def test_ilk_donem_donem_sonu_bakiye(self):
        r = self.s["donemler"]["2025"]["rasyolar"]
        self.assertAlmostEqual(r["Ticari alacak devir hızı"], 17400000 / 3400000)

    def test_gun_sayisi(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(OrderedDict([("2026/09", ORNEK / "mizan_2026.csv")]), Path(tmp) / "r.xlsx", {"2026/09": D(273)})
        r = s["donemler"]["2026/09"]["rasyolar"]
        self.assertAlmostEqual(r["Ortalama tahsil süresi"], 273 / (23200000 / 4500000))


if __name__ == "__main__":
    unittest.main()
