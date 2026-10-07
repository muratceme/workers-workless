import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402


class KodHarfiTesti(unittest.TestCase):
    def test_tablo_1(self):
        # (parti, seviye) -> harf; ISO 2859-1 Tablo 1 ile karşılaştırıldı
        cases = {(8, "II"): "A", (15, "II"): "B", (50, "II"): "D", (281, "I"): "F", (500, "II"): "H",
                 (500, "III"): "J", (501, "II"): "J", (1200, "II"): "J", (1201, "II"): "K", (3200, "S-4"): "G",
                 (10000, "II"): "L", (500001, "III"): "R", (500001, "S-1"): "D"}
        for (parti, seviye), harf in cases.items():
            self.assertEqual(main.kod_harfi(parti, seviye), harf, (parti, seviye))


class TabloTesti(unittest.TestCase):
    def test_satir_basina_gecerli_plan_sayisi(self):
        # Tablo II-A (AQL 0,010-10): her örneklem satırındaki kabul/ret hücresi sayısı
        beklenen = {"A": 1, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6, "H": 7, "J": 8,
                    "K": 9, "L": 9, "M": 9, "N": 9, "P": 9, "Q": 9, "R": 8}
        for harf, adet in beklenen.items():
            r = main.HARFLER.index(harf)
            hucreler = [main._hucre(r, c) for c in range(16)]
            self.assertEqual(sum(isinstance(h, tuple) for h in hucreler), adet, harf)

    def test_bilinen_planlar(self):
        cases = [  # (parti, AQL) -> (örneklem, Ac, Re), seviye II
            (1200, "2.5", (80, 5, 6)), (1200, "4.0", (80, 7, 8)), (1200, "1.0", (80, 2, 3)), (1200, "1.5", (80, 3, 4)),
            (500, "2.5", (50, 3, 4)), (500, "4.0", (50, 5, 6)), (3000, "0.65", (125, 2, 3)),
            (5000, "1.5", (200, 7, 8)),
            (10000, "0.10", (125, 0, 1)),   # L (200) hücresi ↑ → K (125) 0/1
        ]
        for parti, aql, (n, ac, re) in cases:
            p = main.plan_bul(parti, aql)
            self.assertEqual((p.orneklem, p.kabul, p.ret), (n, ac, re), (parti, aql))

    def test_ok_kurallari(self):
        p = main.plan_bul(300, "0.65")              # H (50) hücresi ↓ → J (80) 1/2
        self.assertEqual((p.kod_harfi, p.kullanilan_harf, p.orneklem, p.kabul, p.ret), ("H", "J", 80, 1, 2))
        p = main.plan_bul(100, "1.0")               # F (20) hücresi ↑ → E (13) 0/1
        self.assertEqual((p.kullanilan_harf, p.orneklem, p.kabul), ("E", 13, 0))
        p = main.plan_bul(10, "1.0")                # örneklem ≥ parti → %100 muayene
        self.assertTrue(p.yuzde_yuz)
        self.assertEqual(p.orneklem, 10)

    def test_desteklenmeyen_aql(self):
        with self.assertRaises(ValueError):
            main.plan_bul(1000, "3.0")


class UctanUcaTest(unittest.TestCase):
    def test_ornek_liste(self):
        varsayilan = {"seviye": "II", "kritik": D(0), "major": D("2.5"), "minor": D("4.0")}
        sonuclar = {s["ref"]: s for s in (main.degerlendir(k, varsayilan) for k in
                                         main.liste_oku(KLASOR / "ornek_veri" / "muayeneler.csv"))}
        # 1200 → J (80): majör 2,5 → 5/6, minör 4,0 → 7/8
        self.assertEqual(sonuclar["PO-2026-1041"]["genel"], "KABUL")   # majör 4 ≤ 5, minör 6 ≤ 7
        self.assertEqual(sonuclar["PO-2026-1042"]["kararlar"]["major"], "RET")   # 6 ≥ 6
        self.assertEqual(sonuclar["PO-2026-1043"]["kararlar"]["kritik"], "RET")  # sıfır tolerans
        self.assertEqual(sonuclar["PO-2026-1044"]["kararlar"]["minor"], "RET")   # 450 → H (50): 5/6, 12 hata
        # 60 → E (13): majör 2,5 hücresi ↓ → F (20) 1/2; minör 4,0 → E 1/2
        p = sonuclar["PO-2026-1045"]["planlar"]
        self.assertEqual((p["major"].orneklem, p["major"].kabul), (20, 1))
        self.assertEqual((p["minor"].orneklem, p["minor"].kabul), (13, 1))
        self.assertEqual(sonuclar["PO-2026-1045"]["genel"], "KABUL")
        with tempfile.TemporaryDirectory() as tmp:
            main.rapor_yaz(list(sonuclar.values()), Path(tmp) / "a.xlsx")


if __name__ == "__main__":
    unittest.main()
