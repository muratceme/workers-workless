import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


def kucuk_sinav(tmp: Path, **kw):
    (tmp / "a.csv").write_text("Kitapçık;1;2;3\nA;A;B;C\n", encoding="utf-8")
    (tmp / "c.csv").write_text("Öğrenci No;Sınıf;Kitapçık;1;2;3\ns1;X;A;A;B;C\ns2;X;A;A;B;D\ns3;X;A;A;C;C\n"
                               "s4;Y;A;B;B;\ns5;Y;A;A;D;D\ns6;Y;A;C;C;A\n", encoding="utf-8")
    return main.calistir(tmp / "c.csv", tmp / "a.csv", tmp / "r.xlsx", **kw)


class ElleHesapTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as t:
            cls.s = kucuk_sinav(Path(t))
        cls.m, cls.i = cls.s["maddeler"], cls.s["istat"]
        cls.o = {x["no"]: x for x in cls.s["ogrenciler"]}

    def test_guclik_ve_ayirt_edicilik(self):
        self.assertEqual([round(v["p"], 4) for v in self.m.values()], [round(4 / 6, 4), 0.5, round(2 / 6, 4)])
        # 6 × %27 ≈ 2 kişilik gruplar: üst s1, s2 · alt s5, s6
        self.assertEqual([round(v["d"], 3) for v in self.m.values()], [0.5, 1.0, 0.5])

    def test_kr20(self):
        self.assertAlmostEqual(self.i["kr20"], 1.5 * (1 - (8 / 36 + 9 / 36 + 8 / 36) / (5.5 / 6)), places=9)
        self.assertAlmostEqual(self.i["kr20"], 0.363636, places=5)

    def test_net_ve_sira(self):
        s4 = self.o["s4"]
        self.assertEqual((s4["d"], s4["y"], s4["b"]), (1, 1, 1))
        self.assertAlmostEqual(s4["net"], 0.75)
        self.assertAlmostEqual(s4["puan"], 25.0)
        self.assertEqual(self.o["s1"]["sira"], 1)
        self.assertEqual(self.o["s4"]["sinif_sira"], 1)                              # Y sınıfında en yüksek net

    def test_yanlis_goturmez(self):
        with tempfile.TemporaryDirectory() as t:
            s = kucuk_sinav(Path(t), yanlis_katsayi=0)
        self.assertEqual({x["no"]: x["net"] for x in s["ogrenciler"]}["s4"], 1)


class OrnekSinavTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as t:
            cls.s = main.calistir(ORNEK / "cevaplar.csv", ORNEK / "anahtar.csv", Path(t) / "r.xlsx")

    def test_kitapcik_eslestirme_ve_iptal(self):
        self.assertNotIn(17, self.s["maddeler"])                                      # iptal edilen soru analiz dışı
        self.assertEqual(self.s["istat"]["k"], 19)
        b = next(o for o in self.s["ogrenciler"] if o["kitapcik"] == "B")
        self.assertEqual(set(b["cevap"]), set(range(1, 21)))                         # B cevapları A numaralarına çevrildi

    def test_cok_kolay_madde(self):
        m5 = self.s["maddeler"][5]
        self.assertGreater(m5["p"], 0.9)
        self.assertIn("ayırt etmiyor", m5["yorum"])

    def test_iptal_dogru(self):
        with tempfile.TemporaryDirectory() as t:
            s = main.calistir(ORNEK / "cevaplar.csv", ORNEK / "anahtar.csv", Path(t) / "r.xlsx", iptal_dogru=True)
        o1 = {o["no"]: o for o in self.s["ogrenciler"]}
        o2 = {o["no"]: o for o in s["ogrenciler"]}
        self.assertTrue(all(o2[k]["d"] == o1[k]["d"] + 1 for k in o1))


if __name__ == "__main__":
    unittest.main()
