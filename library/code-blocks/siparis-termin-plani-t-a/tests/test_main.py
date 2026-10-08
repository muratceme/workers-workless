import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class TakvimTesti(unittest.TestCase):
    def test_geri_ve_tatil(self):
        t = main.Takvim(6)
        self.assertEqual(t.geri(date(2026, 11, 2), 1), date(2026, 10, 31))          # Pzt → Cmt (Pazar atlanır)
        self.assertEqual(t.geri(date(2026, 10, 30), 1), date(2026, 10, 28))         # 29 Ekim tatil
        self.assertEqual(main.Takvim(5).geri(date(2026, 11, 2), 1), date(2026, 10, 30))
        self.assertEqual(t.fark(date(2026, 10, 28), date(2026, 10, 30)), 1)
        self.assertEqual(t.ileri(date(2026, 10, 28), 2), date(2026, 10, 30))


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "siparisler.csv", KLASOR / "ta_sablonu.csv", Path(tmp) / "a.xlsx",
                                  ORNEK / "gerceklesen.csv", date(2026, 10, 8))
        cls.x = {x["siparis"]["no"]: x for x in cls.s["sonuc"]}

    def test_geriye_planlama(self):
        p = self.x["PO-503"]["plan"]
        self.assertEqual(p["Sevk"]["bitis"], date(2026, 10, 16))
        self.assertEqual(p["Final Kontrol"]["bitis"], date(2026, 10, 15))
        self.assertEqual(p["Dikim"]["sure"], 4)                                       # 3000 / 900 → 4 gün
        self.assertEqual(p["Kesim"]["bitis"], date(2026, 10, 7))
        self.assertEqual(self.x["PO-502"]["plan"]["Dikim"]["sure"], 10)               # siparişe özel kapasite 600

    def test_durum_ve_tahmin(self):
        a = self.x["PO-503"]
        self.assertTrue(a["plan"]["Kesim"]["durum"].startswith("GECİKTİ"))
        self.assertEqual(a["tahmini_gecikme"], 2)                                     # kumaş 2 gün geç → zincir 2 gün kayar
        self.assertEqual(a["risk"], "Yüksek")
        self.assertEqual((self.x["PO-502"]["pay"], self.x["PO-502"]["tahmini_gecikme"]), (-3, 3))
        self.assertEqual((self.x["PO-501"]["risk"], self.x["PO-501"]["tahmini_gecikme"]), ("Düşük", 0))

    def test_sablon_hatalari(self):
        with tempfile.TemporaryDirectory() as tmp:
            y = Path(tmp) / "s.csv"
            y.write_text("Aşama;Süre;Sonraki Aşama\nA;1;B\nB;1;A\n", encoding="utf-8")
            with self.assertRaises(SystemExit):
                main.sablon_oku(y)
            y.write_text("Aşama;Süre;Sonraki Aşama;Süre Türü\nA;;B;adet\nB;0;\n", encoding="utf-8")
            with self.assertRaises(SystemExit):
                main.sablon_oku(y)                                                    # kapasitesiz adet aşaması


if __name__ == "__main__":
    unittest.main()
