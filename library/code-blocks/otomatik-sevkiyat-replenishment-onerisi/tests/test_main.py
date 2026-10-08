import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


def kur(t: Path, depo_m: int):
    (t / "s.csv").write_text("Mağaza;Ürün Kodu;Stok;Yoldaki\nA;U;2;0\nB;U;10;0\nC;U;5;0\n", encoding="utf-8")
    satir = ["Tarih;Mağaza;Ürün Kodu;Adet"]
    for g in range(1, 29):
        satir.append(f"{g:02d}.09.2026;A;U;1")                   # A: 1/gün
        if g % 2 == 0:
            satir.append(f"{g:02d}.09.2026;B;U;1")               # B: 0,5/gün
    (t / "x.csv").write_text("\n".join(satir) + "\n", encoding="utf-8")
    (t / "d.csv").write_text(f"Ürün Kodu;Stok\nU;{depo_m}\n", encoding="utf-8")
    (t / "p.csv").write_text("Ürün Kodu;Paket\nU;2\n", encoding="utf-8")
    s = main.calistir(t / "s.csv", t / "x.csv", t / "d.csv", t / "r.xlsx", hedef_gun=14, sevk_suresi=2, satis_gunu=28,
                      asgari=2, paket_yolu=t / "p.csv")
    return {x["magaza"]: x for x in s["satirlar"]}, s


class ElleHesapTesti(unittest.TestCase):
    def test_hedef_ve_oneri(self):
        with tempfile.TemporaryDirectory() as t:
            r, _ = kur(Path(t), 100)
        self.assertAlmostEqual(r["A"]["hedef"], 16)                     # 1 × (14 + 2)
        self.assertEqual(r["A"]["oneri"], 14)                           # 16 − 2 = 14 (paket 2'nin katı)
        self.assertEqual(r["B"]["oneri"], 0)                            # hedef 8, stok 10
        self.assertEqual(r["C"]["oneri"], 0)                            # satış yok, asgari 2 < stok 5
        self.assertIn("son dönemde satış yok", r["C"]["not"])
        self.assertEqual(r["A"]["sevk"], 14)

    def test_depo_yetersiz_oncelik(self):
        with tempfile.TemporaryDirectory() as t:
            t = Path(t)
            (t / "s.csv").write_text("Mağaza;Ürün Kodu;Stok\nA;U;0\nB;U;6\n", encoding="utf-8")
            satir = ["Tarih;Mağaza;Ürün Kodu;Adet"] + [f"{g:02d}.09.2026;{m};U;1" for g in range(1, 29) for m in "AB"]
            (t / "x.csv").write_text("\n".join(satir) + "\n", encoding="utf-8")
            (t / "d.csv").write_text("Ürün Kodu;Stok\nU;10\n", encoding="utf-8")
            s = main.calistir(t / "s.csv", t / "x.csv", t / "d.csv", t / "r.xlsx", 14, 2, 28, 2)
        r = {x["magaza"]: x for x in s["satirlar"]}
        self.assertEqual(r["A"]["sevk"] + r["B"]["sevk"], 10)
        self.assertEqual((r["A"]["sevk"], r["B"]["sevk"]), (8, 2))      # önce stoksuz A, eşitlenince sırayla
        self.assertTrue(s["uyarilar"])


class OrnekTest(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as t:
            s = main.calistir(ORNEK / "magaza_stok.csv", ORNEK / "satislar.csv", ORNEK / "depo_stok.csv", Path(t) / "r.xlsx",
                              paket_yolu=ORNEK / "paketler.csv")
        sevk = {}
        for x in s["satirlar"]:
            sevk[x["urun"]] = sevk.get(x["urun"], 0) + x["sevk"]
            if x["urun"] in ("TS-1001-M", "TS-1001-L", "AKS-4004"):
                self.assertEqual(x["sevk"] % (6 if x["urun"].startswith("TS") else 5), 0)
        self.assertEqual(sevk["PNT-3003-38"], 0)
        self.assertLessEqual(sevk["TS-1001-L"], 60)


if __name__ == "__main__":
    unittest.main()
