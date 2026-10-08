import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri" / "taylor_ashe_kumulatif.csv"


class MackTesti(unittest.TestCase):
    """Taylor & Ashe (1983) üçgeni; beklenen değerler Mack (1993) / R ChainLadder MackChainLadder(GenIns) çıktısı."""

    @classmethod
    def setUpClass(cls):
        e, C = main.ucgen_oku(ORNEK, False)
        cls.s = main.hesapla(e, C)

    def test_toplamlar(self):
        self.assertEqual(round(self.s["toplam_rezerv"]), 18680856)
        self.assertEqual(round(self.s["toplam_se"]), 2447095)

    def test_kaza_yili_bazinda(self):
        self.assertEqual([round(r) for r in self.s["rezerv"]],
                         [0, 94634, 469511, 709638, 984889, 1419459, 2177641, 3920301, 4278972, 4625811])
        self.assertEqual(round(self.s["se"][-1]), 1363155)
        self.assertEqual(round(self.s["se"][1]), 75535)

    def test_faktorler(self):
        self.assertEqual([round(x, 4) for x in self.s["f"]],
                         [3.4906, 1.7473, 1.4574, 1.1739, 1.1038, 1.0863, 1.0539, 1.0766, 1.0177])

    def test_artimli_giris_ayni_sonuc(self):
        e, C = main.ucgen_oku(ORNEK, False)
        artimli = [[r[0]] + [r[k] - r[k - 1] for k in range(1, len(r))] for r in C]
        with tempfile.TemporaryDirectory() as tmp:
            yol = Path(tmp) / "a.csv"
            yol.write_text("Kaza;" + ";".join(str(k) for k in range(1, 11)) + "\n" +
                           "\n".join(f"{e[i]};" + ";".join(str(v) for v in r) for i, r in enumerate(artimli)), encoding="utf-8")
            e2, C2 = main.ucgen_oku(yol, True)
            main.rapor(main.hesapla(e2, C2), Path(tmp) / "r.xlsx", "test")
        self.assertEqual(C2, C)

    def test_kuyruk_ve_secili(self):
        e, C = main.ucgen_oku(ORNEK, False)
        s = main.hesapla(e, C, kuyruk=1.05, secili={8: 1.0})
        self.assertEqual(s["f"][8], 1.0)
        self.assertAlmostEqual(s["nihai"][0], C[0][-1] * 1.05)
        self.assertTrue(s["uyarilar"])
        s5 = main.hesapla(e, C, son_n=3)
        self.assertNotEqual(s5["f"][0], self.s["f"][0])


class DokumTesti(unittest.TestCase):
    def test_dokumden_ucgen(self):
        with tempfile.TemporaryDirectory() as tmp:
            yol = Path(tmp) / "d.csv"
            yol.write_text("Hasar No;Kaza Tarihi;Ödeme Tarihi;Tutar\n1;10.03.2024;20.05.2024;100\n2;01.11.2024;15.02.2025;50\n"
                           "3;05.06.2025;10.06.2025;70\n4;05.06.2024;01.01.2026;30\n", encoding="utf-8")
            e, C = main.dokumden_ucgen(yol, "yil", None)
            e2, C2 = main.dokumden_ucgen(yol, "yil", main.tarih("31.12.2025"))
        self.assertEqual(e, ["2024", "2025"])
        self.assertEqual(C, [[100, 150, 180], [70, 70]])
        self.assertEqual(C2, [[100, 150], [70]])


if __name__ == "__main__":
    unittest.main()
