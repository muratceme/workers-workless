import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
KUR = {"USD": D("41.20"), "EUR": D("48.10")}


def calistir(tmp):
    return main.calistir(ORNEK / "recete.csv", ORNEK / "malzeme_maliyetleri.csv", Path(tmp) / "a.xlsx", ORNEK / "rota.csv",
                         ORNEK / "is_merkezleri.csv", ORNEK / "parti_miktarlari.csv", ORNEK / "satis_fiyatlari.csv", KUR)


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as tmp:
            cls.s = calistir(tmp)
        cls.u = cls.s["urunler"]

    def test_yari_mamul_elle_hesap(self):
        a = self.u["YM-AYAK"]
        # kereste 0,012 × 1,15 × 38.500 + tutkal 0,05 × 95
        self.assertAlmostEqual(float(a["malzeme"]), 531.30 + 4.75, places=6)
        # CNC (6 + 60/120) dk, zımpara (5 + 10/120) dk
        self.assertAlmostEqual(float(a["iscilik"]), 6.5 / 60 * 260 + (5 + 10 / 120) / 60 * 210, places=6)
        self.assertAlmostEqual(float(a["gug"]), 6.5 / 60 * 640 + (5 + 10 / 120) / 60 * 150, places=6)

    def test_doviz_ve_ust_seviye(self):
        m = self.u["MSA-01"]
        vernik = 0.60 * 1.05 * 9.40 * 48.10
        vida = 16 * 1.02 * 0.55
        self.assertAlmostEqual(float(m["malzeme"]), 2 * float(self.u["YM-AYAK"]["malzeme"]) + 1150 + vida + vernik + 135, places=6)
        # montaj 2 kişi: (20 + 30/20) dk × 220 × 2
        montaj = [o for o in m["ops"] if o["im"] == "MONTAJ"][0]
        self.assertAlmostEqual(float(montaj["iscilik"]), 21.5 / 60 * 220 * 2, places=6)
        self.assertEqual(self.s["mamuller"], ["MSA-01", "SND-01"])

    def test_toplam_ve_marj(self):
        for k in ("SND-01", "MSA-01", "YM-OTURAK"):
            u = self.u[k]
            self.assertEqual(u["toplam"], u["malzeme"] + u["iscilik"] + u["gug"])
        s = self.u["SND-01"]
        self.assertAlmostEqual(float(s["marj"]), float((D("2150") - s["toplam"]) / D("2150") * 100), places=6)
        self.assertFalse(self.s["uyarilar"])

    def test_agac_toplami_malzemeye_esit(self):
        pay = main.malzeme_payi(self.u, "SND-01")
        self.assertAlmostEqual(float(sum(pay.values())), float(self.u["SND-01"]["malzeme"]), places=6)


class KenarDurumTesti(unittest.TestCase):
    def test_dongu_eksik_maliyet_gug_orani(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "r.csv").write_text("Ana Ürün;Bileşen;Miktar;Çıktı Miktarı\nA;B;2;\nA;X;1;\nB;C;10;5\n", encoding="utf-8")
            (t / "m.csv").write_text("Malzeme Kodu;Birim Maliyet;Para Birimi\nC;3;\nX;1;GBP\n", encoding="utf-8")
            (t / "ro.csv").write_text("Ürün Kodu;İş Merkezi;Süre\nA;W;30\n", encoding="utf-8")
            (t / "im.csv").write_text("İş Merkezi;İşçilik\nW;100\n", encoding="utf-8")
            s = main.calistir(t / "r.csv", t / "m.csv", t / "a.xlsx", t / "ro.csv", t / "im.csv", gug_orani=50)
            u = s["urunler"]
            self.assertEqual(u["B"]["malzeme"], D(6))                 # 10 C → 5 B: B başına 2 × 3
            self.assertEqual(u["A"]["malzeme"], D(12))                # 2 B; X kuru yok → 0
            self.assertEqual((u["A"]["iscilik"], u["A"]["gug"]), (D(50), D(25)))
            self.assertTrue(any("GBP" in x for x in s["uyarilar"]))
            (t / "r.csv").write_text("Ana Ürün;Bileşen;Miktar\nA;B;1\nB;A;1\n", encoding="utf-8")
            with self.assertRaises(SystemExit):
                main.calistir(t / "r.csv", t / "m.csv", t / "a.xlsx")

    def test_kur(self):
        self.assertEqual(main.kur_coz(["usd=41,20"]), {"USD": D("41.20")})


if __name__ == "__main__":
    unittest.main()
