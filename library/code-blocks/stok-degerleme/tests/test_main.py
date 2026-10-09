import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal as D
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

ORNEK = main.ORNEK


def r2(x):
    return x.quantize(D("0.01"))


class OrnekVeriTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.cikti = Path(cls.tmp.name) / "s.xlsx"
        cls.s = main.calistir(ORNEK / "hareketler.csv", cls.cikti, "fifo", None, ORNEK / "satis_fiyatlari.csv")
        cls.k = {s.kod: s.sonuc for s in cls.s["stoklar"]}

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_uc_yontem(self):                         # elle hesaplanmış değerler
        v = self.k["STK-01"]
        self.assertEqual((v["fifo"]["deger"], v["fifo"]["smm"]), (D("750.00"), D("8150.00")))
        self.assertEqual((v["hareketli"]["deger"], v["hareketli"]["smm"]), (D("684.00"), D("8216.00")))
        self.assertEqual((v["donemsel"]["deger"], v["donemsel"]["smm"]), (D("667.50"), D("8232.50")))

    def test_iadeler(self):
        v = self.k["STK-04"]                           # satış iadesi son çıkış maliyetiyle, alış iadesi kendi fiyatıyla
        self.assertEqual((r2(v["fifo"]["deger"]), r2(v["fifo"]["smm"])), (D("60457.14"), D("47542.86")))
        self.assertEqual((v["hareketli"]["deger"], v["hareketli"]["smm"]), (D("59840.00"), D("48160.00")))
        self.assertEqual(v["donemsel"]["deger"], D("60000.00"))

    def test_eksi_stok_ve_korunum(self):
        v = self.k["STK-02"]["fifo"]
        self.assertEqual((v["miktar"], v["deger"], v["duzeltme"]), (D("90"), D("30150.00"), D("750.00")))
        for kod, sonuc in self.k.items():
            for y, r in sonuc.items():
                giren = {"STK-01": D("8900"), "STK-02": D("104200"), "STK-03": D("12325"), "STK-04": D("108000")}[kod]
                self.assertEqual(r2(r["deger"] + r["smm"]), giren, (kod, y))

    def test_uyarilar_ve_excel(self):
        u = {(x["tur"], x["stok"]) for x in self.s["uyarilar"]}
        self.assertTrue({("Eksi stok", "STK-02"), ("Hareketsiz stok", "STK-03"), ("Değer düşüklüğü", "STK-03")} <= u)
        self.assertEqual(next(x["dusukluk"] for x in self.s["ngd"] if x["s"].kod == "STK-03"), D("5525.00"))
        wb = load_workbook(self.cikti)
        self.assertEqual(wb.sheetnames, ["Değerleme Özeti", "Yöntem Karşılaştırması", "Hareket Detayı", "FIFO Katmanları", "Değer Düşüklüğü", "Uyarılar"])


class KuralTesti(unittest.TestCase):
    def test_tarih_ve_tur(self):
        self.assertEqual(main.tur_bul("Satış İadesi", None), "satis_iade")
        self.assertEqual(main.tur_bul("Giriş", None), "satis_iade")
        self.assertEqual(main.tur_bul("Sarf", None), "cikis")
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "hareketler.csv", Path(tmp) / "o.xlsx", "hareketli", date(2026, 1, 31))
            v = {x.kod: x.sonuc for x in s["stoklar"]}
            self.assertEqual((v["STK-01"]["hareketli"]["miktar"], v["STK-01"]["hareketli"]["deger"]), (D("1500"), D("3200.00")))
            wb = load_workbook(Path(tmp) / "o.xlsx")
            self.assertNotIn("Değer Düşüklüğü", wb.sheetnames)

    def test_cli(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(main.main(["--cikti", str(Path(tmp) / "x.xlsx"), "--yontem", "donemsel"]), 0)
            self.assertEqual(main.main(["--tarih", "2026"]), 2)
            self.assertEqual(main.main(["--hareketler", str(Path(tmp) / "yok.csv")]), 1)


if __name__ == "__main__":
    unittest.main()
