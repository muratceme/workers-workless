import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class PuanTesti(unittest.TestCase):
    def test_uzunluk_esikleri(self):
        self.assertEqual([main.hata_puani(x, False) for x in (0.5, 3, 3.01, 6, 6.5, 9, 9.01, 40)], [1, 1, 2, 2, 3, 3, 4, 4])

    def test_delik(self):
        self.assertEqual((main.hata_puani(1, True), main.hata_puani(1.2, True)), (2, 4))

    def test_yard_basina_en_fazla_4(self):
        toplam, yardlar = main.top_puani([{"konum_yd": 5.1, "uzunluk_inc": 10, "delik": False},
                                          {"konum_yd": 5.8, "uzunluk_inc": 10, "delik": False},
                                          {"konum_yd": 6.2, "uzunluk_inc": 2, "delik": False}])
        self.assertEqual((toplam, yardlar), (5, {5: 4, 6: 1}))

    def test_surekli_hata(self):
        toplam, yardlar = main.top_puani([{"konum_yd": 10.0, "uzunluk_inc": 36 * 5, "delik": False, "yon": "boy"}])
        self.assertEqual(toplam, 20)                                      # 5 yard × 4
        self.assertEqual(sorted(yardlar), [10, 11, 12, 13, 14])

    def test_formul_yayimlanmis_ornek(self):
        # 120 yd × 45 inç, 22 puan → 22 × 36 × 100 / (45 × 120) = 14,67
        self.assertAlmostEqual(main.puan_100yd2(22, 45, 120), 14.666, places=2)


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "toplar.csv", ORNEK / "hata_kayitlari.csv", Path(tmp) / "k.xlsx")
        cls.t = cls.s["toplar"]

    def test_toplar(self):
        self.assertEqual({k: v["puan"] for k, v in self.t.items()}, {"T-001": 12, "T-002": 133, "T-003": 5, "T-004": 3, "T-005": 5})
        # T-001: 150 cm = 59,06 inç, 100 m = 109,36 yd → 12 × 3600 / (59,06 × 109,36)
        self.assertAlmostEqual(self.t["T-001"]["p100yd2"], 12 * 3600 / ((150 / 2.54) * (100 / 0.9144)))
        self.assertEqual(self.t["T-002"]["sonuc"], "RET")
        self.assertEqual(self.t["T-003"]["sonuc"], "KABUL")

    def test_notlar(self):
        self.assertIn("ardışık", self.t["T-002"]["notlar"][0])
        self.assertIn("en", self.t["T-003"]["notlar"][0])
        self.assertIn("etiketin 5 altında", self.t["T-005"]["notlar"][0])

    def test_parti_alan_agirlikli(self):
        p = self.s["partiler"]["P-101"]
        alan = sum(self.t[k]["boy_yd"] * self.t[k]["en_inc"] / 36 for k in ("T-001", "T-002", "T-003"))
        self.assertAlmostEqual(p["p100yd2"], 150 * 100 / alan)
        self.assertEqual((p["sonuc"], p["ret"]), ("KABUL", 1))

    def test_inc_birimi(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "t.csv").write_text("Top No;Uzunluk;Kesilebilir En\nA;120;45\n", encoding="utf-8")
            (t / "h.csv").write_text("Top No;Konum;Puan\n" + "".join(f"A;{i};1\n" for i in range(22)), encoding="utf-8")
            s = main.calistir(t / "t.csv", t / "h.csv", t / "k.xlsx", birim="inc")
        self.assertAlmostEqual(s["toplar"]["A"]["p100yd2"], 14.666, places=2)


if __name__ == "__main__":
    unittest.main()
