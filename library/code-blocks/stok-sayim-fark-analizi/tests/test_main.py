import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class IpucuTesti(unittest.TestCase):
    def test_sifir_ve_rakam(self):
        self.assertIn("Fazla/eksik sıfır", main.tek_satir_ipuclari(D(30), D(300), None)[0])
        self.assertIn("yer değiştirmiş", main.tek_satir_ipuclari(D(54), D(45), None)[0])
        self.assertEqual(main.tek_satir_ipuclari(D(54), D(54), None), [])
        self.assertEqual(main.tek_satir_ipuclari(D(50), D(41), None), [])     # fark 9 ama rakamlar farklı

    def test_koli(self):
        self.assertIn("Koli/adet", main.tek_satir_ipuclari(D(5), D(120), D(24))[0])
        self.assertIn("tam katı", main.tek_satir_ipuclari(D(800), D(700), D(100))[0])

    def test_kod_benzer(self):
        self.assertTrue(main.kod_benzer("BJ-1000-KIR", "BJ-1000-MAV"))
        self.assertFalse(main.kod_benzer("VD-M8", "KB-2.5"))

    def test_baslik_katlama(self):
        self.assertEqual(main.bul(["LOKASYON", "STOK KODU", "SAYILAN MİKTAR"], "sayılan miktar"), 2)
        self.assertEqual(main.bul(["Stok Kodu", "Sayilan Miktar"], "sayılan miktar"), 1)


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "sayim.csv", ORNEK / "sistem_stogu.csv", Path(tmp) / "a.xlsx",
                                  ORNEK / "bekleyen_hareketler.csv")
        cls.r = {(r["lok"], r["kod"]): r for r in cls.s["satirlar"]}

    def test_tutarlar(self):
        o = self.s["ozet"]
        self.assertEqual(o["eksik_tutar"], D("-21029"))
        self.assertEqual(o["fazla_tutar"], D("27000"))
        # KB-2.5 iki lokasyonda ±500: ürün bazında mahsup edilir
        self.assertEqual(o["urun_eksik"], D("-11829"))
        self.assertEqual(o["urun_fazla"], D("17800"))
        self.assertEqual(o["satir"], 19)
        self.assertEqual(o["dogru"], 7)

    def test_sayim_satirlari_toplanir_ve_bekleyen(self):
        self.assertEqual(self.r[("A-01", "VD-M8")]["fark"], 0)               # 700 + 500
        bnt = self.r[("A-02", "BNT-48")]
        self.assertEqual((bnt["sistem"], bnt["fark"]), (D(180), 0))         # 200 − 20 çıkış

    def test_ipuclari_ve_ikinci_sayim(self):
        kb = self.r[("B-01", "KB-2.5")]
        self.assertIn("ürün toplamında fark yok", kb["ipuclari"][0])
        self.assertIn("hiç sayılmamış", kb["ikinci_sayim"])
        self.assertIn("varyant", self.r[("A-02", "BJ-1000-KIR")]["ipuclari"][0])
        self.assertIn("negatif", self.r[("B-01", "PN-PNS")]["ipuclari"][0])
        self.assertEqual(self.r[("A-02", "XX-9081")]["ikinci_sayim"], "kartı yok")
        self.assertEqual(self.r[("A-01", "MS-MSK")]["ikinci_sayim"], "")    # %0,4 ve 26 TL
        self.assertEqual(self.r[("B-01", "MT-DRL")]["tutar"], D("-10350"))

    def test_tolerans_ve_lokasyonsuz(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "say.csv").write_text("Stok Kodu;Sayılan\nA;99\nA;0\nB;10\n", encoding="utf-8")
            (t / "sis.csv").write_text("Depo;Stok Kodu;Miktar;Birim Maliyet\nD1;A;60;2\nD2;A;40;2\nD1;B;10;5\n", encoding="utf-8")
            (t / "bek.csv").write_text("Stok Kodu;Miktar\nB;-1\n", encoding="utf-8")
            s = main.calistir(t / "say.csv", t / "sis.csv", t / "a.xlsx", t / "bek.csv", tolerans=1.0)
        r = {x["kod"]: x for x in s["satirlar"]}
        self.assertEqual((r["A"]["sistem"], r["A"]["fark"], r["A"]["tolerans_ici"]), (D(100), D(-1), True))
        self.assertEqual((r["B"]["sistem"], r["B"]["fark"], r["B"]["tolerans_ici"]), (D(9), D(1), False))


if __name__ == "__main__":
    unittest.main()
