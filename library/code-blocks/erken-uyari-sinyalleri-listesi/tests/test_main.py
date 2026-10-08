import json
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


def ornek(t, **k):
    return main.calistir(ORNEK / "portfoy.csv", Path(t) / "r.xlsx", ORNEK / "kkb_cek_senet.csv", ORNEK / "haciz.csv",
                         ORNEK / "ciro.csv", date(2026, 9, 30), **k)


class OrnekTest(unittest.TestCase):
    def test_puan_ve_sinif(self):
        with tempfile.TemporaryDirectory() as t:
            s = ornek(t)
        f = {x["musteri"]: x for x in s["firmalar"]}
        self.assertEqual((f["1001"]["puan"], f["1001"]["sinif"]), (100, "Kritik"))      # 25+15+30+15+20 → 100 tavan
        self.assertEqual((f["1003"]["puan"], f["1003"]["sinif"]), (87, "Kritik"))       # 60 + kamu 15 + özel 12
        self.assertEqual((f["1006"]["puan"], f["1006"]["sinif"]), (52, "Yakın İzleme"))  # 25 + 15 + 12 (2025 çeki pencere dışı)
        self.assertEqual((f["1002"]["puan"], f["1002"]["sinif"]), (28, "İzleme"))       # doluluk 5 + senet 10+3 + ciro 10
        self.assertEqual(f["1007"]["puan"], 10)                                          # geçen yıla göre %40
        self.assertEqual(f["1005"]["puan"], 0)
        self.assertFalse(f["9999"]["portfoyde"])
        self.assertEqual(s["firmalar"][0]["musteri"], "1001")

    def test_ciro_kiyas(self):
        with tempfile.TemporaryDirectory() as t:
            s = ornek(t)
        c = {x["anahtar"]: x["aciklama"] for x in s["sinyaller"] if x["tur"] == "Ciro düşüşü"}
        self.assertIn("geçen yılın", c["1001"])
        self.assertIn("önceki 3 aya", c["1002"])
        self.assertIn("%55", c["1001"])

    def test_ayar(self):
        with tempfile.TemporaryDirectory() as t:
            (Path(t) / "a.json").write_text(json.dumps({"gecikme": [[1, 5], [91, 80]]}), encoding="utf-8")
            s = ornek(t, ayar_yolu=Path(t) / "a.json")
        f = {x["musteri"]: x for x in s["firmalar"]}
        self.assertEqual(f["1004"]["puan"], 5)
        self.assertEqual(f["1003"]["puan"], 100)


class YardimciTest(unittest.TestCase):
    def test_donem(self):
        self.assertEqual(main.donem("09.2026"), (2026, 9))
        self.assertEqual(main.donem("2026-09"), (2026, 9))
        self.assertEqual(main.donem("15.09.2026"), (2026, 9))
        self.assertEqual(main.ay_kaydir((2026, 2), -3), (2025, 11))

    def test_para(self):
        self.assertEqual(main.para("2.500"), 2500)
        self.assertEqual(main.para("1.234.567,50"), main.Decimal("1234567.50"))


if __name__ == "__main__":
    unittest.main()
