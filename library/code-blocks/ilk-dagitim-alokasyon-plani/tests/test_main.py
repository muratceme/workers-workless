import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class YardimciTesti(unittest.TestCase):
    def test_en_buyuk_kalan(self):
        self.assertEqual(main.en_buyuk_kalan(10, {"a": 1, "b": 1, "c": 1}), {"a": 4, "b": 3, "c": 3})
        self.assertEqual(sum(main.en_buyuk_kalan(97, {"a": 3.3, "b": 1.7, "c": 0.9}).values()), 97)
        self.assertEqual(main.en_buyuk_kalan(0, {"a": 1}), {"a": 0})

    def test_beden_payi_normallesir(self):
        egri = {"A": {"S": 20, "M": 40, "L": 40, "36": 50}}
        p = main.beden_payi("m1", "A", ["S", "M", "L"], egri, {})
        self.assertAlmostEqual(sum(p.values()), 1.0)
        self.assertAlmostEqual(p["S"], 0.2)
        self.assertEqual(main.beden_payi("m1", "A", ["38", "40"], egri, {"38": 0.5, "40": 0.5}), {"38": 0.5, "40": 0.5})


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "depo_stok.csv", ORNEK / "magazalar.csv", Path(tmp) / "a.xlsx", ORNEK / "beden_egrisi.csv")
        cls.p = cls.s["plan"]

    def test_rezerv_ve_toplam(self):
        # Kapasite sınırı olmayan bir beden: TS-1001 Beyaz M 260 × 0,7 = 182 dağıtılır (KON-Merkez azaltması hariç)
        dagitilan = sum(v for k, v in self.p.items() if k[1:] == ("TS-1001", "Beyaz", "M"))
        self.assertLessEqual(dagitilan, 182)
        self.assertGreaterEqual(dagitilan, 175)

    def test_pasif_magaza_ve_kapasite(self):
        self.assertNotIn("ESK-Odunpazarı", {k[0] for k in self.p})
        self.assertEqual(sum(v for k, v in self.p.items() if k[0] == "KON-Merkez"), 50)
        self.assertTrue(any("KON-Merkez" in u for u in self.s["uyarilar"]))

    def test_asgari_ve_buyuk_magaza_daha_fazla(self):
        xl = {k[0]: v for k, v in self.p.items() if k[1:] == ("ELB-2002", "Lacivert", "XL")}
        self.assertTrue(all(v >= 1 for m, v in xl.items() if m != "KON-Merkez"))
        m = lambda mg: self.p.get((mg, "TS-1001", "Siyah", "M"), 0)  # noqa: E731
        self.assertGreater(m("İST-Nişantaşı"), m("SAM-Atakum"))

    def test_kirik_seri_ve_asorti(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "s.csv").write_text("Model;Renk;S;M;L;XL\nX;Kırmızı;10;3;10;10\n", encoding="utf-8")
            (t / "m.csv").write_text("Mağaza;Küme;Potansiyel\nm1;A;5\nm2;A;4\nm3;B;3\nm4;C;2\nm5;C;1\n", encoding="utf-8")
            s = main.calistir(t / "s.csv", t / "m.csv", t / "a.xlsx", rezerv=0)
            self.assertEqual({z["magaza"] for z in s["kirik"]}, {"m4", "m5"})     # M yalnız 3 adet: m1-m3'e
            a = main.calistir(t / "s.csv", t / "m.csv", t / "b.xlsx", rezerv=0, asorti={"S": 1, "M": 1, "L": 1, "XL": 1})
            self.assertEqual(sum(a["plan"].values()), 12)                          # 3 paket × 4
            self.assertEqual(a["plan"][("m1", "X", "Kırmızı", "M")], 1)


if __name__ == "__main__":
    unittest.main()
