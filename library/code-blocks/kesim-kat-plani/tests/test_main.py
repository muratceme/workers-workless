import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


def kes(pastallar):
    k = {}
    for p in pastallar:
        for b, v in p["oran"].items():
            k[b] = k.get(b, 0) + v * p["kat"]
    return k


class PlanTesti(unittest.TestCase):
    def test_basit(self):
        p = main.planla({"S": 10, "M": 10}, max_kat=10, max_urun=2, min_kat=1)
        self.assertEqual([(x["kat"], x["oran"]) for x in p], [(10, {"S": 1, "M": 1})])

    def test_oran_kalanla_orantili(self):
        self.assertEqual(main.oran_sec({"S": 80, "M": 160, "L": 160, "XL": 80}, 80, 6), {"S": 1, "M": 2, "L": 2, "XL": 1})

    def test_son_pastal_en_az_fazla(self):
        fazla, kat, oran = main.son_pastal({"S": 9, "M": 18}, max_kat=20, max_urun=3, min_kat=5)
        self.assertEqual((fazla, kat, oran), (0, 9, {"S": 1, "M": 2}))


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as t:
            cls.s = main.calistir(ORNEK / "siparis.csv", Path(t) / "a.xlsx", 80, 6, 5, 2.0, ORNEK / "tuketim.csv", 3.0)

    def test_kisitlar_ve_kapsam(self):
        for renk, x in self.s.items():
            for p in x["pastallar"]:
                self.assertLessEqual(sum(p["oran"].values()), 6, renk)
                self.assertTrue(5 <= p["kat"] <= 80, (renk, p))
            k = kes(x["pastallar"])
            for b, a in x["hedef"].items():
                self.assertGreaterEqual(k.get(b, 0), a, (renk, b))               # hedef her bedende karşılanır
            fazla = sum(k.values()) - sum(x["hedef"].values())
            self.assertLessEqual(fazla, max(6, round(0.02 * sum(x["hedef"].values()))))

    def test_lacivert_tam_kesim(self):
        x = self.s["TS-101 Lacivert"]
        self.assertEqual(kes(x["pastallar"]), dict(x["hedef"]))
        self.assertEqual(len(x["pastallar"]), 3)

    def test_kumas(self):
        p = self.s["TS-101 Beyaz"]["pastallar"][0]
        boy = 0.62 + 2 * 0.66 + 2 * 0.70 + 0.75 + 0.06                         # S + 2M + 2L + XL + uç payları
        self.assertAlmostEqual(p["boy"], boy)
        self.assertAlmostEqual(p["kumas"], boy * 80)


if __name__ == "__main__":
    unittest.main()
