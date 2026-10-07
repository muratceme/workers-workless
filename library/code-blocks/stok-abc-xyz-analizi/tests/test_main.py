import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class SiniflamaTesti(unittest.TestCase):
    def test_abc_siniri_asan_kalem_ust_sinifta(self):
        s = main.abc_sinifla({"K1": D(70), "K2": D(20), "K3": D(6), "K4": D(4)}, 80, 95)
        # K1 öncesi %0 → A; K2 öncesi %70 → A; K3 öncesi %90 → B; K4 öncesi %96 → C
        self.assertEqual({k: v[0] for k, v in s.items()}, {"K1": "A", "K2": "A", "K3": "B", "K4": "C"})
        self.assertAlmostEqual(s["K3"][1], 96.0)

    def test_xyz(self):
        self.assertEqual(main.xyz_sinifla([10.0] * 6, 0.5, 1.0), ("X", 0.0))
        sinif, cv = main.xyz_sinifla([20, 0, 20, 0, 20], 0.5, 1.0)       # ort 12, σ = √96
        self.assertEqual(sinif, "Y")
        self.assertAlmostEqual(cv, 96 ** 0.5 / 12)
        self.assertEqual(main.xyz_sinifla([0, 0, 100, 0, 0, 0], 0.5, 1.0)[0], "Z")

    def test_donem(self):
        from datetime import date
        for x in ("2026-03", "01.03.2026", "03.2026", "Mart 2026", date(2026, 3, 15)):
            self.assertEqual(main.donem_anahtari(x), "2026-03")


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "aylik_tuketim.csv", Path(tmp) / "a.xlsx", ORNEK / "birim_maliyetler.csv")
        cls.k = cls.s["kalemler"]

    def test_siniflar(self):
        beklenen = {"HM-001": "AX", "YM-010": "AX", "HM-002": "AY", "YM-012": "BX", "YM-011": "BX", "SR-030": "BX",
                    "HM-003": "CZ", "YD-021": "CX", "YD-022": "CZ", "YD-020": "CZ", "SR-031": "C?", "ES-099": "Hareketsiz"}
        self.assertEqual({k: v["sinif"] for k, v in self.k.items()}, beklenen)

    def test_deger_ve_yeni_kalem(self):
        self.assertEqual(self.k["HM-001"]["deger"], D("3006250.0"))     # 48.100 kg × 62,50
        self.assertEqual(self.k["SR-031"]["etkin_donem"], 4)            # ilk tüketimden önceki 8 ay sayılmaz
        self.assertEqual(self.k["YD-022"]["etkin_donem"], 8)

    def test_genis_bicim(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "g.csv").write_text("Stok Kodu;Birim Maliyet;2026-01;2026-02;2026-03\nA;10;5;5;5\nB;1;0;9;0\n", encoding="utf-8")
            s = main.calistir(t / "g.csv", t / "a.xlsx", en_az_donem=1)
        self.assertEqual(s["donemler"], ["2026-01", "2026-02", "2026-03"])
        self.assertEqual(s["kalemler"]["A"]["sinif"], "AX")
        self.assertEqual(s["kalemler"]["B"]["xyz"], "Y")                 # [9, 0]: CV = 1,0 → sınır dahil Y


if __name__ == "__main__":
    unittest.main()
