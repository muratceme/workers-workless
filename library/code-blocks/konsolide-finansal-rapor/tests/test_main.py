import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

O = KLASOR / "ornek_veri"


def ornek(t, **k):
    return main.calistir(O / "sirketler.csv", Path(t) / "r.xlsx", O / "eliminasyonlar.csv", **k)


class OrnekTest(unittest.TestCase):
    def test_elle_hesap(self):
        with tempfile.TemporaryDirectory() as t:
            s = ornek(t)
        t_ = s["tablo"]
        # Aktif: 7.320.000 + 2.830.000 + 1.760.000 − 600.000 − 150.000 (cari) − 90.000 (stok kârı) − 1.800.000 (yatırım) + 200.000 (şerefiye)
        self.assertEqual(t_["aktif_top"], Decimal("9470000"))
        self.assertEqual(t_["pasif_top"], Decimal("9470000"))
        # Kâr: 420.000 + 400.000 + 300.000 − 90.000; KGO %20 × 300.000
        self.assertEqual(t_["donem_net"], Decimal("1030000"))
        self.assertEqual(s["kgo_kar"], Decimal("60000"))
        self.assertEqual(s["ana_payi"], Decimal("970000"))
        self.assertEqual(s["serefiye"], Decimal("200000"))                 # 800.000 − %80 × 750.000
        self.assertEqual(s["kgo"], Decimal("240000"))                      # %20 × (900.000 + 300.000)
        k = s["konsolide"]
        self.assertEqual(k["570"]["net"], Decimal("-1120000"))             # 800.000 + 200.000 (BP) + %80 × 150.000 (BL)
        self.assertEqual(k["245"]["net"], 0)
        self.assertEqual(k["500"]["net"], Decimal("-4000000"))
        self.assertEqual(k["600"]["net"], Decimal("-8500000"))             # 10.800.000 − 2.000.000 − 300.000
        self.assertEqual(k["320"]["net"], Decimal("-1410000"))             # 50.000 mutabakatsız kalır

    def test_bulgular(self):
        with tempfile.TemporaryDirectory() as t:
            s = ornek(t)
        self.assertEqual([(b["seviye"], b["kontrol"]) for b in s["bulgular"]], [("Hata", "Mutabakat farkı")])
        self.assertEqual(s["bulgular"][0]["tutar"], Decimal("-50000"))

    def test_kayitlar_denk(self):
        with tempfile.TemporaryDirectory() as t:
            s = ornek(t)
        for no in {r["no"] for r in s["kayitlar"]}:
            r = [x for x in s["kayitlar"] if x["no"] == no]
            self.assertEqual(sum(x["borc"] for x in r), sum(x["alacak"] for x in r), no)

    def test_eliminasyonsuz(self):
        with tempfile.TemporaryDirectory() as t:
            s = main.calistir(O / "sirketler.csv", Path(t) / "r.xlsx")
        self.assertEqual(s["tablo"]["aktif_top"], s["tablo"]["pasif_top"])
        self.assertEqual(s["tablo"]["donem_net"], Decimal("1120000"))


class ParaTest(unittest.TestCase):
    def test_binlik(self):
        self.assertEqual(main.para("750.000"), Decimal("750000"))
        self.assertEqual(main.para("1.000.000,50"), Decimal("1000000.50"))


if __name__ == "__main__":
    unittest.main()
