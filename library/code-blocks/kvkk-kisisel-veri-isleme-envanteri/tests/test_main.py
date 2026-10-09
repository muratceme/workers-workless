import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

ORNEK = main.ORNEK


class OrnekVeriTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.cikti = Path(cls.tmp.name) / "k.xlsx"
        cls.s = main.calistir(ORNEK / "departman_formlari.csv", cls.cikti, ORNEK / "departmanlar.csv")
        cls.b = {(b["no"], b["tur"]) for b in cls.s["bulgular"]}
        cls.k = {k.no: k for k in cls.s["kayitlar"]}

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_ozel_nitelikli(self):
        self.assertEqual(self.k[3].ozel_turler, ["Sağlık"])
        self.assertEqual(self.k[13].ozel_turler, ["Biyometrik"])
        self.assertEqual(self.k[5].ozel_turler, ["Ceza mahkûmiyeti"])
        self.assertTrue({(3, "Özel nitelikli veri sebebi"), (13, "Özel nitelikli veri sebebi"), (3, "Özel nitelik işareti"), (13, "Biyometrik veri")} <= self.b)
        self.assertNotIn((5, "Özel nitelikli veri sebebi"), self.b)                 # kanunlarda açıkça öngörülme (md. 6/3-a)

    def test_yurtdisi_ve_diger(self):
        self.assertTrue({(9, "Yurt dışı aktarım dayanağı"), (8, "Yurt dışı aktarımda açık rıza"),
                         (11, "Standart sözleşme bildirimi"), (8, "Ölçülülük"), (4, "Belirsiz saklama süresi"), (11, "Belirsiz saklama süresi"),
                         (7, "Gereksiz açık rıza"), (15, "Mükerrer kayıt"), ("", "Form gelmedi")} <= self.b)
        self.assertNotIn((1, "Gereksiz açık rıza"), self.b)

    def test_sebep_ve_kelimeler(self):
        b, t = main.sebep_coz("Sözleşmenin ifası, hukuki yükümlülük; uydurma")
        self.assertEqual([m for _, m in b], ["5/2-c", "5/2-ç"])
        self.assertEqual(t, ["uydurma"])
        self.assertEqual(main.ozel_bul("Reklam ve tanıtım", "Satış raporu"), [])
        self.assertEqual(main.ozel_bul("", "Tanı kodu, din"), ["Din / inanç", "Sağlık"])

    def test_excel(self):
        wb = load_workbook(self.cikti)
        self.assertEqual(wb.sheetnames, ["Bulgular", "Envanter", "Özel Nitelikli Veriler", "Yurt Dışı Aktarımlar", "Saklama ve İmha", "Kategori Özeti",
                                         "Departman Durumu"])
        self.assertEqual(wb["Envanter"].max_row, 15)                                 # mükerrer kamera kaydı bir kez
        d = {r[0]: r[7] for r in wb["Departman Durumu"].iter_rows(min_row=2, values_only=True)}
        self.assertEqual((d["Üretim"], d["İnsan Kaynakları"]), ("Form gelmedi", "Düzeltme gerekli"))

    def test_cli(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(main.main(["--cikti", str(Path(tmp) / "x.xlsx")]), 0)
            self.assertEqual(main.main(["--formlar", str(Path(tmp) / "yok.csv")]), 1)


if __name__ == "__main__":
    unittest.main()
