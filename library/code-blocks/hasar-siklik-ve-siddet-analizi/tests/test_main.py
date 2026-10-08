import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


def kucuk_portfoy(t: Path, **kw):
    (t / "p.csv").write_text("Poliçe No;Başlangıç;Bitiş;Prim;Tip\nA;01.01.2025;01.01.2026;1000;X\nB;01.07.2025;01.07.2026;2000;Y\n",
                             encoding="utf-8")
    (t / "h.csv").write_text("Hasar No;Poliçe No;Hasar Tarihi;Ödenen;Muallak;Durum\n1;A;10.03.2025;4000;1000;Açık\n"
                             "2;B;05.08.2025;3000;0;Kapalı\n3;B;05.02.2026;1000;0;Kapalı\n4;A;01.06.2025;9000;0;Red\n", encoding="utf-8")
    return main.calistir(t / "p.csv", t / "h.csv", t / "r.xlsx", ["Tip"], degerleme=date(2025, 12, 31), **kw)


class ElleHesapTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as t:
            cls.s = kucuk_portfoy(Path(t), buyuk_hasar=None)
        cls.y = cls.s["donem"]["2025"]

    def test_maruziyet_ve_prim(self):
        self.assertAlmostEqual(self.y["maruziyet"], 1 + 184 / 365)                    # B: 01.07–31.12 = 184 gün
        self.assertAlmostEqual(self.y["prim"], 1000 + 2000 * 184 / 365)
        self.assertNotIn("2026", self.s["donem"])                                      # değerleme sonrası kazanılmamış

    def test_siklik_siddet(self):
        e = 1 + 184 / 365
        self.assertEqual(self.y["adet"], 2)                                             # Red ve değerleme sonrası hariç
        self.assertAlmostEqual(self.y["siklik"], 2 / e)
        self.assertAlmostEqual(self.y["siddet"], 4000)                                  # (5000 + 3000) / 2
        self.assertAlmostEqual(self.y["hp_orani"], 8000 / (1000 + 2000 * 184 / 365))

    def test_segment(self):
        x = self.s["segment"]["Tip"]
        self.assertEqual((x["X"]["adet"], x["Y"]["adet"]), (1, 1))
        self.assertAlmostEqual(x["Y"]["maruziyet"], 184 / 365)

    def test_ceyrek_ve_buyuk_hasar(self):
        with tempfile.TemporaryDirectory() as t:
            s = kucuk_portfoy(Path(t), tur="ceyrek", buyuk_hasar="4000")
        self.assertAlmostEqual(s["donem"]["2025-Ç1"]["maruziyet"], 90 / 365)
        self.assertAlmostEqual(s["toplam"]["siddet_sinirli"], (4000 + 3000) / 2)
        self.assertEqual(s["toplam"]["buyuk"], 1)

    def test_yuzdelik_esik(self):
        self.assertAlmostEqual(main.esik_hesapla([10, 20, 30, 40, 50], "p50"), 30)


class OrnekTest(unittest.TestCase):
    def test_ornek_calisir(self):
        with tempfile.TemporaryDirectory() as t:
            s = main.calistir(ORNEK / "policeler.csv", ORNEK / "hasarlar.csv", Path(t) / "r.xlsx", ["Araç Tipi", "Bölge"],
                              degerleme=date(2026, 9, 30))
        self.assertEqual(sorted(s["donem"]), ["2024", "2025", "2026"])
        self.assertGreater(s["segment"]["Araç Tipi"]["Kamyon"]["siklik"], s["segment"]["Araç Tipi"]["Otomobil"]["siklik"])


if __name__ == "__main__":
    unittest.main()
