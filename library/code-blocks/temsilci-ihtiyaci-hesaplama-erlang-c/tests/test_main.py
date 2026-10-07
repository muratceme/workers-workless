import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402


class ErlangCTesti(unittest.TestCase):
    """Yayımlanmış çözümlü örnek (callcentrehelper.com, Erlang C): 100 çağrı / 30 dk, AHT 180 sn, 14 temsilci
    → trafik 10, P(bekleme) 0,1741, SL(20 sn) 0,88835, ASA 7,84 sn, doluluk 0,7143"""

    def test_yayimlanmis_ornek(self):
        s = main.hesapla(100, 30, 180, n=14)
        self.assertAlmostEqual(s.trafik, 10.0)
        self.assertAlmostEqual(s.bekleme_olasiligi, 0.1741319, places=6)
        self.assertAlmostEqual(s.servis_seviyesi, 0.88835, places=5)
        self.assertAlmostEqual(s.ort_cevap_suresi, 7.84, places=2)
        self.assertAlmostEqual(s.doluluk, 0.7142857, places=6)

    def test_en_kucuk_temsilci(self):
        s = main.hesapla(100, 30, 180, sl_hedef=0.80, hedef_sn=20, maks_doluluk=1.0)
        self.assertEqual(s.temsilci, 13 if main.hesapla(100, 30, 180, n=13).servis_seviyesi >= 0.80 else 14)
        self.assertGreaterEqual(s.servis_seviyesi, 0.80)
        self.assertLess(main.hesapla(100, 30, 180, n=s.temsilci - 1).servis_seviyesi, 0.80)

    def test_doluluk_siniri_ve_kayip(self):
        s = main.hesapla(100, 30, 180, maks_doluluk=0.70, kayip=0.30)
        self.assertLessEqual(s.doluluk, 0.70)
        self.assertEqual(s.planlanan, -(-s.temsilci * 10 // 7))       # tavan(n / 0,7)

    def test_buyuk_trafik_kararli(self):
        s = main.hesapla(3000, 60, 300)                                 # 250 Erlang
        self.assertGreater(s.temsilci, 250)
        self.assertGreaterEqual(s.servis_seviyesi, 0.80)

    def test_sifir_cagri(self):
        self.assertEqual(main.hesapla(0, 30, 180).temsilci, 0)


class UctanUcaTest(unittest.TestCase):
    def test_ornek(self):
        tahmin = main.tahmin_oku(KLASOR / "ornek_veri" / "cagri_tahmini.csv")
        satirlar = [(t, main.hesapla(t["cagri"], 30, 180)) for t in tahmin]
        self.assertEqual(len(satirlar), 22)
        with tempfile.TemporaryDirectory() as tmp:
            main.rapor_yaz(satirlar, Path(tmp) / "e.xlsx", {"aralik": 30, "sl": 0.8, "hedef_sn": 20, "maks_doluluk": 0.85, "kayip": 0.3})


if __name__ == "__main__":
    unittest.main()
