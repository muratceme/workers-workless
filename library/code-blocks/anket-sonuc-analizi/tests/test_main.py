import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

ORNEK = main.ORNEK


class IstatistikTesti(unittest.TestCase):
    def test_ki_kare_p_bilinen_degerler(self):
        # Ki-kare kritik değer tablosu: p = 0,05
        for ki, sd in ((3.841, 1), (5.991, 2), (7.815, 3), (16.919, 9), (31.410, 20)):
            self.assertAlmostEqual(main.ki_kare_p(ki, sd), 0.05, places=4)
        self.assertAlmostEqual(main.ki_kare_p(6.635, 1), 0.01, places=4)
        self.assertEqual(main.ki_kare_p(0, 3), 1.0)

    def test_ki_kare_tablo(self):
        # 2×2: [[20, 30], [30, 20]] → beklenen 25; ki = 4 × (5² / 25) = 4
        t = main.ki_kare([[20, 30], [30, 20]])
        self.assertAlmostEqual(t["ki"], 4.0)
        self.assertEqual(t["sd"], 1)
        self.assertAlmostEqual(t["p"], 0.0455, places=4)
        self.assertAlmostEqual(t["cramer_v"], 0.2)
        self.assertTrue(t["guvenilir"])
        self.assertFalse(main.ki_kare([[2, 3], [1, 2]])["guvenilir"])       # beklenen < 5
        self.assertIsNone(main.ki_kare([[5, 0], [7, 0]]))                   # tek sütun dolu

    def test_hata_payi_ve_kodlama(self):
        self.assertAlmostEqual(main.hata_payi(0.5, 400), 0.049, places=3)
        kodlar = [{"soru": "S10", "kod": "Fiyat", "kelimeler": ["fiyat", "pahali"]}, {"soru": "*", "kod": "Ambalaj", "kelimeler": ["ambalaj"]}]
        self.assertEqual(main.acik_kodla("Fiyatı PAHALI, ambalajını sevmedim", kodlar, "S10"), ["Fiyat", "Ambalaj"])
        self.assertEqual(main.acik_kodla("Uyumsuz fiyatlandırma", kodlar, "S11"), [])        # kod yalnız S10 için
        self.assertEqual(main.acik_kodla("Şeffaf bir politika", kodlar, "S10"), [])          # "fiyat" kelime başında değil
        self.assertEqual(main.maskele("bana 0532 111 22 33 veya a@b.com yazın"), "bana [TELEFON] veya [E-POSTA] yazın")


class OrnekVeriTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.cikti = Path(cls.tmp.name) / "a.xlsx"
        cls.s = main.calistir(ORNEK / "yanitlar.csv", ORNEK / "sorular.csv", cls.cikti, ORNEK / "kod_cercevesi.csv")

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_frekans_ve_olcek(self):
        s = self.s
        self.assertEqual(s["n"], 240)                                         # mükerrer Y006 satırı alınmadı
        f = [x for x in s["frekans"] if x["soru"].kod == "S1"]
        self.assertEqual(f[0]["yanitlayan"], 238)                             # iki boş yanıt
        self.assertEqual(sum(x["n"] for x in f), 238)
        self.assertAlmostEqual(sum(x["oran"] for x in f), 1.0)
        coklu = [x for x in s["frekans"] if x["soru"].kod == "S2"]
        self.assertGreater(sum(x["oran"] for x in coklu), 1.0)                # çoklu seçim %100'ü aşar
        o = {x["soru"].kod: x for x in s["olcek"]}
        self.assertEqual(o["S3"]["n"], 239)                                   # '7' geçersiz
        self.assertAlmostEqual(o["S3"]["ust2"] + o["S3"]["alt2"] + o["S3"]["dagilim"][3] / 239, 1.0)

    def test_capraz_ve_uyarilar(self):
        s = self.s
        anlamli = {(x["soru"].kod, x["kirilim"].kod) for x in s["anlamli"]}
        self.assertIn(("S1", "Yaş Grubu"), anlamli)                           # veride yaşa bağlı kurgulandı
        c = next(x for x in s["capraz"] if x["soru"].kod == "S1" and x["kirilim"].kod == "Yaş Grubu")
        self.assertEqual(sum(c["payda"]), 238)
        turler = {(u["tur"], u["kim"]) for u in s["uyarilar"]}
        self.assertTrue({("Geçersiz değer", "S3"), ("Geçersiz değer", "S8"), ("Mükerrer yanıt no", "Y006")} <= turler)
        self.assertTrue({"Y010", "Y058", "Y121", "Y189"} <= set(s["duz"]))
        a = s["acik"]["S10"]
        self.assertLessEqual(len(a["kodsuz"]), a["n"])
        self.assertGreater(a["kodlar"]["Fiyat"], 0)
        self.assertTrue(any("[TELEFON]" in y for _, y in a["kodsuz"]))
        self.assertFalse(any("0532" in y for _, y in a["kodsuz"]))
        wb = load_workbook(self.cikti)
        self.assertEqual(wb.sheetnames, ["Özet", "Frekanslar", "Ölçek Soruları", "Çapraz Tablolar", "Anlamlı Farklar", "Açık Uçlu Kodlar",
                                         "Kodlanmayan Yanıtlar", "Uyarılar"])

    def test_cli(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(main.main(["--cikti", str(Path(tmp) / "x.xlsx"), "--kirilim", "Cinsiyet", "Olmayan"]), 0)
            self.assertEqual(main.main(["--yanitlar", str(Path(tmp) / "yok.csv")]), 1)


if __name__ == "__main__":
    unittest.main()
