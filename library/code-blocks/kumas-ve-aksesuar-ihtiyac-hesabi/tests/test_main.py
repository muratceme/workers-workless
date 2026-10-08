import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "siparis.csv", ORNEK / "recete.csv", Path(tmp) / "a.xlsx",
                                  ORNEK / "malzeme_karti.csv", ORNEK / "stok.csv", 3.0)
        cls.i = cls.s["ihtiyac"]

    def test_genis_siparis(self):
        self.assertEqual(self.s["toplam_adet"], D(3400))
        self.assertEqual(len(self.s["siparis"]), 12)

    def test_kumas_elle_hesap(self):
        x = self.i[("KMS-SUP-160", "Beyaz", "-")]
        self.assertEqual(x["brut"], D("1500") * D("1.03") * D("0.22") * D("1.08"))   # 367,092 kg
        self.assertEqual(x["net"], x["brut"] - 120)
        self.assertEqual(x["top"], 10)                                                 # 247,09 / 25
        self.assertAlmostEqual(float(x["metre"]), float(x["brut"] / (D("1.80") * D("0.160"))), places=6)

    def test_bedene_gore_tuketim(self):
        # SW-201: S 150 × 0,50 + M 250 × 0,55 + L 250 × 0,60 + XL 100 × 0,65 = 427,5 kg (firesiz, fazla kesimsiz)
        self.assertEqual(self.i[("KMS-3IP-320", "Antrasit", "-")]["net_tuketim"], D("427.50"))

    def test_aksesuar_yuvarlama(self):
        self.assertEqual(self.i[("ETK-BEDEN", "-", "S")]["siparis"], D(1000))        # 682,9 → 500'ün katı
        self.assertEqual(self.i[("ETK-MARKA", "-", "-")]["siparis"], D(3000))        # 2.772 → 3.000
        self.assertEqual(self.i[("POSET", "-", "-")]["siparis"], D(2100))
        self.assertEqual(self.i[("IPL-40-2", "Beyaz", "-")]["siparis"], D(295000))   # 5.000 m bobin katı
        self.assertEqual(self.s["uyarilar"], [])

    def test_uzun_bicim_ve_asgari(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "s.csv").write_text("Model;Renk;Beden;Adet\nA;Kırmızı;M;10\nB;Mavi;M;5\n", encoding="utf-8")
            (t / "r.csv").write_text("Model;Malzeme Kodu;Tür;Birim;Tüketim;Renge Bağlı\nA;K1;Kumaş;m;1,5;evet\nA;D1;Aksesuar;adet;6;\n",
                                     encoding="utf-8")
            (t / "k.csv").write_text("Malzeme Kodu;Ambalaj Miktarı;Asgari Sipariş\nD1;144;288\n", encoding="utf-8")
            s = main.calistir(t / "s.csv", t / "r.csv", t / "a.xlsx", t / "k.csv")
        self.assertEqual(s["ihtiyac"][("K1", "Kırmızı", "-")]["brut"], D("15.0"))
        self.assertEqual(s["ihtiyac"][("D1", "-", "-")]["siparis"], D(288))           # 60 → 144 → asgari 288
        self.assertTrue(any("B" in u for u in s["uyarilar"]))                          # reçetesiz model


if __name__ == "__main__":
    unittest.main()
