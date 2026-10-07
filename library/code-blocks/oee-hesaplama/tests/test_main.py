import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402


class KlasikOrnekTesti(unittest.TestCase):
    """OEE literatüründeki klasik örnek (oee.com): 480 dk vardiya, 60 dk mola, 47 dk duruş,
    1,0 sn ideal çevrim, 19.271 adet, 423 hatalı → K %88,81 · P %86,11 · Q %97,80 · OEE %74,79"""

    def test_klasik(self):
        o, uyarilar = main.satir_olcum({"vardiya_suresi": 480, "planli_durus": 60, "plansiz_durus": 47,
                                        "cevrim": 1.0, "toplam": 19271, "hatali": 423})
        self.assertAlmostEqual(o.kullanilabilirlik, 0.8881, places=4)
        self.assertAlmostEqual(o.performans, 0.8611, places=4)
        self.assertAlmostEqual(o.kalite, 0.9780, places=4)
        self.assertAlmostEqual(o.oee, 0.7479, places=4)
        self.assertEqual(uyarilar, [])


class KontrolTesti(unittest.TestCase):
    def test_performans_yuzu_asamaz_uyarisi(self):
        _, u = main.satir_olcum({"vardiya_suresi": 480, "planli_durus": 0, "plansiz_durus": 0,
                                 "cevrim": 10, "toplam": 5000, "hatali": 0})
        self.assertTrue(any("gerçekçi değil" in x for x in u))

    def test_negatif_calisma(self):
        o, u = main.satir_olcum({"vardiya_suresi": 480, "planli_durus": 60, "plansiz_durus": 500,
                                 "cevrim": 10, "toplam": 1, "hatali": 0})
        self.assertIsNone(o)

    def test_toplama_oran_ortalamasi_degil(self):
        a = main.Olcum(100, 100, 100, 10, 10)    # OEE %100, kısa vardiya
        b = main.Olcum(400, 200, 100, 10, 10)    # OEE %25, uzun vardiya
        self.assertAlmostEqual((a + b).oee, 200 / 500)   # ortalama (%62,5) değil, %40


class UctanUcaTest(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(KLASOR / "ornek_veri" / "vardiya_kayitlari.csv", Path(tmp) / "o.xlsx")
            self.assertEqual(list(s["makineler"]), ["Enjeksiyon-01", "Enjeksiyon-02", "CNC-07"])
            self.assertTrue(0 < s["genel"].oee < 1)


if __name__ == "__main__":
    unittest.main()
