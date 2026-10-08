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


class YardimciTesti(unittest.TestCase):
    def test_hucre(self):
        self.assertEqual(main.hucre_coz("Ç+2")[:2], ("Ç", D(2)))
        self.assertEqual(main.hucre_coz("x 2,5")[:2], ("Ç", D("2.5")))
        self.assertEqual(main.hucre_coz(1)[:2], ("Ç", D(0)))
        self.assertEqual(main.hucre_coz("ht")[0], "HT")
        self.assertEqual(main.hucre_coz("YI")[0], "Yİ")
        self.assertEqual(main.hucre_coz("HT+2")[0], "?")              # fazla mesai yalnız çalışılan güne
        self.assertEqual(main.hucre_coz("Ö")[0], "?")
        self.assertIsNone(main.hucre_coz("")[0])

    def test_tatiller(self):
        tam, yarim = main.tatiller(2026)
        self.assertIn(date(2026, 5, 27), tam)
        self.assertIn(date(2026, 5, 26), yarim)
        self.assertIn(date(2026, 10, 28), yarim)
        tam27, _ = main.tatiller(2027)
        self.assertIn("Atatürk", tam27[date(2027, 5, 19)])           # Kurban 4. günü 19 Mayıs'la çakışır

    def test_gunler(self):
        self.assertEqual(main._gunler([1, 2, 3, 7, 9, 10]), "1-3, 7, 9-10")

    def test_donem(self):
        self.assertEqual(main.donem_bul([["Firma — Mayıs 2026 Puantaj"]]), (2026, 5))
        self.assertEqual(main.donem_bul([["Dönem: 2026-09"]]), (2026, 9))


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "puantaj_mayis_2026.csv", Path(tmp) / "a.xlsx")
        cls.b = {(x["sicil"], x["kontrol"]): x for x in cls.s["bulgular"]}
        cls.o = {x["sicil"]: x for x in cls.s["ozetler"]}

    def test_donem_basliktan(self):
        self.assertEqual(self.s["donem"], (2026, 5))

    def test_yasal(self):
        self.assertEqual(self.b[("1002", "Günlük 11 saat")]["gunler"], "7")
        self.assertIn("274", self.b[("1002", "Yıllık 270 saat")]["aciklama"])
        self.assertEqual(self.b[("1005", "Hafta tatili")]["gunler"], "4-10")
        self.assertNotIn(("1001", "Hafta tatili"), self.b)

    def test_tatil_kurallari(self):
        self.assertEqual(self.b[("1002", "Genel tatilde çalışma")]["gunler"], "19")
        self.assertIn("arife", self.b[("1004", "GT kodu")]["aciklama"])
        self.assertEqual(self.b[("1001", "İzinde genel tatil")]["gunler"], "19")
        self.assertEqual(self.b[("1003", "Devamsızlık ve hafta tatili")]["gunler"], "21, 24")

    def test_sgk_ve_toplamlar(self):
        self.assertEqual((self.o["1003"]["sgk"], self.o["1003"]["eksik_kodu"]), (26, "12"))   # 3 rapor + 1 devamsızlık
        self.assertEqual((self.o["1004"]["istihdam"], self.o["1004"]["sgk"]), (21, 21))
        self.assertEqual((self.o["1006"]["sgk"], self.o["1006"]["eksik_kodu"]), (17, "21"))
        self.assertEqual(self.o["1001"]["sgk"], 30)
        self.assertIn(("1003", "Toplam: çalışılan gün"), self.b)
        self.assertIn(("1005", "Toplam: fazla mesai"), self.b)
        self.assertIn(("1006", "İstihdam dışı gün"), self.b)
        self.assertFalse(any(k[1] == "Toplam: SGK gün" for k in self.b))

    def test_bos_ve_ayda_olmayan_gun(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            gun = ";".join(str(d) for d in range(1, 32))
            hucre = ["Ç"] * 30 + ["Ç"]
            hucre[4] = ""
            (t / "p.csv").write_text(f"Sicil;Ad Soyad;{gun};SGK Gün\n1;A;{';'.join(hucre)};31\n", encoding="utf-8")
            s = main.calistir(t / "p.csv", t / "a.xlsx", donem=(2026, 2))
        b = {x["kontrol"]: x for x in s["bulgular"]}
        self.assertEqual(b["Boş gün"]["gunler"], "5")
        self.assertIn("Ayda olmayan gün", b)
        self.assertEqual(b["Toplam: SGK gün"]["seviye"], "Hata")


if __name__ == "__main__":
    unittest.main()
