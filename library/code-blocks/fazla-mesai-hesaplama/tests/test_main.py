import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class KuralTesti(unittest.TestCase):
    def test_ara_dinlenme(self):
        self.assertEqual([main.ara_dinlenme(D(x)) for x in ("4", "4.5", "7.5", "8", "12")],
                         [D("0.25"), D("0.5"), D("0.5"), D(1), D(1)])


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "pdks.csv", Path(tmp) / "a.xlsx", ORNEK / "personel.csv")
        cls.o = {x["sicil"]: x for x in cls.s["ozet"]}

    def test_fazla_calisma_ve_ucret(self):
        ali = self.o["1001"]
        # Cumartesi 08-13, mola yok → 5 sa − 30 dk = 4,5 sa; her hafta 49,5 → 4,5 sa FÇ × 3 hafta
        self.assertEqual(ali["fm"], D("13.5"))
        self.assertEqual(ali["fm_ucret"], D("4050.00"))                 # 13,5 × 45.000/225 × 1,5
        self.assertEqual(ali["ubgt_ucret"], D("1500.00"))               # 19 Mayıs: 45.000/30
        self.assertEqual(ali["serbest"], D("20.25"))                    # 13,5 × 1,5

    def test_fazla_surelerle(self):
        ayse = self.o["1002"]
        self.assertEqual((ayse["fm"], ayse["fsc"]), (D(0), D("7.5")))   # 40 saatlik sözleşme: 3 + 4,5
        self.assertAlmostEqual(float(ayse["fsc_ucret"]), 7.5 * 38000 / 225 * 1.25, places=1)

    def test_gece_vardiyasi_ve_sinirlar(self):
        meh = self.o["1003"]
        self.assertEqual(meh["fm"], D(9))                                # 22:00–07:00, 60 dk mola = 8 sa × 6 gece
        self.assertTrue(meh["yillik_asim"])                              # 262 + 9 = 271
        turler = {(i["sicil"], i["tur"]) for i in self.s["gunluk_ihlal"]}
        self.assertIn(("1002", "11 saat"), turler)
        self.assertIn(("1003", "gece 7,5 saat"), turler)
        self.assertNotIn(("1001", "gece 7,5 saat"), turler)

    def test_denklestirme_ve_ubgt_haric(self):
        with tempfile.TemporaryDirectory() as tmp:
            # Hafta 1: 50 sa, hafta 2: 40 sa → denkleştirmeyle 2 haftada ortalama 45, fazla çalışma yok
            satir = ["Sicil;Tarih;Çalışma Saati"]
            for g in range(4, 9):
                satir.append(f"1;{g:02d}.05.2026;10")
            for g in range(11, 16):
                satir.append(f"1;{g:02d}.05.2026;8")
            (Path(tmp) / "k.csv").write_text("\n".join(satir) + "\n", encoding="utf-8")
            s0 = main.calistir(Path(tmp) / "k.csv", Path(tmp) / "a.xlsx")
            s2 = main.calistir(Path(tmp) / "k.csv", Path(tmp) / "b.xlsx", denklestirme=2)
            s3 = main.calistir(ORNEK / "pdks.csv", Path(tmp) / "c.xlsx", ORNEK / "personel.csv", ubgt_haric=True)
        self.assertEqual(s0["ozet"][0]["fm"], D(5))
        self.assertEqual(s2["ozet"][0]["fm"], D(0))
        ali = {x["sicil"]: x for x in s3["ozet"]}["1001"]
        self.assertEqual(ali["fm"], D(9))                                 # 3. hafta: 49,5 − 9 (19 Mayıs) = 40,5


if __name__ == "__main__":
    unittest.main()
