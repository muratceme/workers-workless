import json
import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
AYAR = json.loads((KLASOR / "senaryolar.json").read_text(encoding="utf-8"))["senaryolar"]


def islem(no, gun, tutar, yon=1, kanal="nakit", karsi="", saat=10):
    return main.Islem(no, datetime(2026, 9, 1, saat) + timedelta(days=gun), "M", "", kanal, yon, D(tutar), karsi, "", "")


class SenaryoTesti(unittest.TestCase):
    def test_parcali(self):
        a = []
        main.s1_parcali([islem("1", 0, 95000), islem("2", 1, 90000), islem("3", 2, 99000), islem("4", 9, 85000)], AYAR["S1"], a)
        self.assertEqual([x["islemler"] for x in a], [["1", "2", "3"]])
        a = []
        main.s1_parcali([islem("1", 0, 95000), islem("2", 1, 100000), islem("3", 2, 99000)], AYAR["S1"], a)
        self.assertEqual(a, [])                                    # 100.000 eşiğin altı değil

    def test_hizli_gecis(self):
        a = []
        main.s2_hizli_gecis([islem("1", 0, 300000, kanal="transfer"), islem("2", 1, 280000, yon=-1, kanal="transfer")], AYAR["S2"], a)
        self.assertEqual(len(a), 1)
        a = []
        main.s2_hizli_gecis([islem("1", 0, 300000, kanal="transfer"), islem("2", 5, 280000, yon=-1, kanal="transfer")], AYAR["S2"], a)
        self.assertEqual(a, [])                                    # 48 saat dışında

    def test_cok_gonderen(self):
        a = []
        isl = [islem(str(i), 0, 10000, kanal="transfer", karsi=f"Kişi {i}", saat=i) for i in range(10)]
        main.s5_cok_gonderen(isl, AYAR["S5"], a)
        self.assertEqual(len(a), 1)
        a = []
        main.s5_cok_gonderen(isl[:9] + [islem("x", 0, 10000, kanal="transfer", karsi="KİŞİ 1", saat=20)], AYAR["S5"], a)
        self.assertEqual(a, [])                                    # aynı kişi büyük harfle → 9 farklı

    def test_kelime_tam_kelime(self):
        a = []
        x = islem("1", 0, 1000)
        x.aciklama = "Alfabet kırtasiye"                           # "bet" kelime içinde → sayılmaz
        y = islem("2", 0, 1000)
        y.aciklama = "BET ödeme"
        main.s8_kelime([x, y], AYAR["S8"], a)
        self.assertEqual([z["islemler"] for z in a], [["2"]])


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = main.calistir(ORNEK / "islemler.csv", KLASOR / "senaryolar.json", Path(tmp) / "a.xlsx",
                                  ORNEK / "musteriler.csv", ORNEK / "riskli_ulkeler.csv", {"USD": D("41.20")})
        cls.m = {k["musteri"]: k for k in cls.s["musteriler"]}

    def test_musteriler(self):
        self.assertNotIn("M001", self.m)                           # olağan maaş hesabı alarm üretmez
        self.assertEqual({k: sorted(v["senaryolar"]) for k, v in self.m.items()},
                         {"M002": ["S1", "S2", "S6"], "M003": ["S5", "S6", "S7", "S8", "S9"], "M004": ["S4", "S6"],
                          "M005": ["S6", "S7"], "M006": ["S3", "S9"]})
        self.assertEqual((self.m["M003"]["puan"], self.m["M003"]["oncelik"]), (100, "Yüksek"))
        self.assertEqual((self.m["M006"]["puan"], self.m["M006"]["oncelik"]), (25, "Düşük"))

    def test_doviz(self):
        a = [x for x in self.s["alarmlar"] if x["senaryo"] == "S4"][0]
        self.assertEqual(a["tutar"], D("391400.00"))              # 9.500 USD × 41,20


if __name__ == "__main__":
    unittest.main()
