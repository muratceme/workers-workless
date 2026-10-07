import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class FifoTesti(unittest.TestCase):
    def test_fifo_kapama(self):
        satirlar = [["Cari", "Tarih", "Vade", "Belge No", "Borç", "Alacak"],
                    ["A", "01.01.2026", "31.01.2026", "F1", "100", "0"],
                    ["A", "05.01.2026", "04.02.2026", "F2", "50", "0"],
                    ["A", "10.02.2026", "", "T1", "0", "120"]]
        kalemler, fazla, _ = main.acik_kalemler(satirlar, 0)
        self.assertEqual([(k.belge, k.kalan) for k in kalemler], [("F2", D(30))])   # F1 tamamen, F2'nin 20'si kapandı

    def test_avans(self):
        satirlar = [["Cari", "Tarih", "Belge No", "Borç", "Alacak"],
                    ["A", "01.01.2026", "T0", "0", "200"], ["A", "02.01.2026", "F1", "150", "0"]]
        kalemler, fazla, uyarilar = main.acik_kalemler(satirlar, 30)
        self.assertEqual(kalemler, [])
        self.assertEqual(fazla["A"], D(50))

    def test_dilimler(self):
        self.assertEqual([main.dilim(g) for g in (-5, 0, 1, 30, 31, 90, 91, 180, 181)],
                         ["Vadesi gelmemiş", "Vadesi gelmemiş", "1-30 gün", "1-30 gün", "31-60 gün", "61-90 gün",
                          "91-180 gün", "91-180 gün", "180+ gün"])


class UctanUcaTest(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "cari_hareketler.csv", date(2026, 9, 30), Path(tmp) / "a.xlsx",
                              limit_yolu=ORNEK / "kredi_limitleri.csv")
        m = s["musteri"]
        # Deneme: 120.000 faturası 100.000 havaleyle kısmen kapandı → 20.000 (vade 05.07, 87 gün) + 80.000 (vade 20.07, 72 gün) + 45.000 (vadesi gelmemiş)
        dl = m["Deneme Lojistik Ltd. Şti."]
        self.assertEqual(dl["61-90 gün"], D(100000))
        self.assertEqual(dl["Vadesi gelmemiş"], D(45000))
        self.assertEqual(dl["toplam"], D(145000))
        kt = m["Kurgu Tekstil A.Ş."]
        self.assertEqual(kt["180+ gün"], D(64500))                 # vade 01.04 → 182 gün
        self.assertEqual(kt["1-30 gün"], D(18000))
        self.assertNotIn("Hayali Kırtasiye", m)                    # fazla tahsilat: avans
        self.assertTrue(any("Hayali Kırtasiye" in u for u in s["uyarilar"]))
        self.assertEqual(s["genel"]["toplam"], D(145000 + 82500 + 210000))


if __name__ == "__main__":
    unittest.main()
