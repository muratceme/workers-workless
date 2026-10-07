import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
BAS = date(2026, 11, 2)


def kur(t, mps, bom, stok, acik=None):
    (t / "m.csv").write_text(mps, encoding="utf-8")
    (t / "b.csv").write_text(bom, encoding="utf-8")
    (t / "s.csv").write_text(stok, encoding="utf-8")
    if acik:
        (t / "a.csv").write_text(acik, encoding="utf-8")
    return main.calistir(t / "m.csv", t / "b.csv", t / "s.csv", t / "o.xlsx", BAS, 8, 7, (t / "a.csv") if acik else None)


def dolu(liste):
    return {d: v for d, v in enumerate(liste) if v}


class DersKitabiTesti(unittest.TestCase):
    def test_uc_seviye_l4l(self):
        # A: 5. dönem 100, 8. dönem 150; A→2B, B→3C; TS: A 1, B 2, C 1; eldeki A 20, B 50, C 100
        with tempfile.TemporaryDirectory() as tmp:
            s = kur(Path(tmp), "Ürün;Dönem;Miktar\nA;5;100\nA;8;150\n", "Üst Malzeme;Alt Malzeme;Miktar\nA;B;2\nB;C;3\n",
                    "Malzeme;Eldeki;Tedarik Süresi\nA;20;1\nB;50;2\nC;100;1\n")
        k = s["kayit"]
        self.assertEqual(dolu(k["A"]["satir"]["pveris"]), {4: 80, 7: 150})
        self.assertEqual(dolu(k["B"]["satir"]["brut"]), {4: 160, 7: 300})
        self.assertEqual(dolu(k["B"]["satir"]["pveris"]), {2: 110, 5: 300})
        self.assertEqual(dolu(k["C"]["satir"]["brut"]), {2: 330, 5: 900})
        self.assertEqual(dolu(k["C"]["satir"]["pveris"]), {1: 230, 4: 900})
        self.assertEqual(s["gecmis"], [])

    def test_lot_emniyet_acik_siparis_fire(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = kur(Path(tmp), "Ürün;Dönem;Miktar\nX;2;100\n", "Üst Malzeme;Alt Malzeme;Miktar;Fire %\nX;Y;3;10\n",
                    "Malzeme;Eldeki;Emniyet Stoğu;Tedarik Süresi;Lot Yöntemi;Lot Miktarı;En Az Sipariş;Ambalaj Katı\n"
                    "X;0;0;1;L4L;;;\nY;50;20;0;Sabit;100;;\n")
        y = s["kayit"]["Y"]["satir"]
        self.assertAlmostEqual(y["brut"][1], 330)                       # 100 × 3 × 1,10
        self.assertAlmostEqual(y["net"][1], 300)                        # 50 − 330 = −280 → ES 20'ye çıkmak için 300
        self.assertEqual(y["pgiris"][1], 300)                           # sabit lot 100'ün katı
        self.assertAlmostEqual(y["eldeki"][1], 20)

    def test_moq_ve_gecmis(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = kur(Path(tmp), "Ürün;Dönem;Miktar\nX;1;30\n", "Üst Malzeme;Alt Malzeme;Miktar\n",
                    "Malzeme;Eldeki;Tedarik Süresi;En Az Sipariş;Ambalaj Katı\nX;0;2;100;25\n", )
        self.assertEqual(s["kayit"]["X"]["satir"]["pgiris"][1], 100)
        self.assertEqual(s["gecmis"], [("X", 1, 100, -1)])              # 2 dönem önce verilmeliydi

    def test_dusuk_seviye_kodu(self):
        bom = {"A": [("B", 1, 0), ("C", 1, 0)], "B": [("C", 1, 0)]}
        self.assertEqual(main.dusuk_seviye_kodlari(bom, {"A"}), {"A": 0, "B": 1, "C": 2})
        with self.assertRaises(SystemExit):
            main.dusuk_seviye_kodlari({"A": [("B", 1, 0)], "B": [("A", 1, 0)]}, {"A"})


class OrnekTesti(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "mps.csv", ORNEK / "urun_agaci.csv", ORNEK / "malzemeler.csv", Path(tmp) / "o.xlsx", BAS, 8, 7,
                              ORNEK / "acik_siparisler.csv")
        k = s["kayit"]
        self.assertEqual(k["VDA-M6"]["llc"], 2)                         # hem masada hem ayakta → en alt seviye
        self.assertEqual(k["MDF-18"]["satir"]["giris"][2], 60)          # 09.11 açık sipariş 2. dönem
        for kod, x in k.items():                                       # emniyet stoğu korunur
            for d in range(1, 9):
                self.assertGreaterEqual(round(x["satir"]["eldeki"][d], 6), x["kalem"]["emniyet"], kod)


if __name__ == "__main__":
    unittest.main()
