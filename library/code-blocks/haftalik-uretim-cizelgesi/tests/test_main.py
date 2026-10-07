import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
PZT = date(2026, 11, 2)


def yaz(t, ad, metin):
    (t / ad).write_text(metin, encoding="utf-8")
    return t / ad


class TakvimTesti(unittest.TestCase):
    def test_vardiya_yayilimi(self):
        tk = main.Takvim(PZT, 5, set())
        bas, bit, parca = tk.ilerle(1, datetime(2026, 11, 2, 12, 30), 630)    # 12:30–16:00 = 210, Salı 08:00 + 420
        self.assertEqual((bas, bit), (datetime(2026, 11, 2, 12, 30), datetime(2026, 11, 3, 15, 0)))
        self.assertEqual(parca, [(date(2026, 11, 2), 210), (date(2026, 11, 3), 420)])

    def test_hafta_sonu_ve_tatil(self):
        tk = main.Takvim(PZT, 5, {date(2026, 11, 9)})
        _, bit, _ = tk.ilerle(1, datetime(2026, 11, 6, 15, 0), 120)           # Cuma 60 dk, Cmt-Paz-Pzt(tatil) atlanır, Salı 60 dk
        self.assertEqual(bit, datetime(2026, 11, 10, 9, 0))

    def test_uc_vardiya_gece(self):
        tk = main.Takvim(PZT, 5, set())
        _, bit, parca = tk.ilerle(3, datetime(2026, 11, 2, 8, 0), 1500)       # 24 saat kesintisiz: 1440 + Salı 60
        self.assertEqual(bit, datetime(2026, 11, 3, 9, 0))


class CizelgeTesti(unittest.TestCase):
    def test_elle_hesaplanan(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            s = yaz(t, "s.csv", "Sipariş No;Ürün;Miktar;Termin\nA;P1;40;03.11.2026\nB;P1;100;05.11.2026\nC;P2;10;05.11.2026\n")
            r = yaz(t, "r.csv", "Ürün;Operasyon Sırası;Operasyon;İş Merkezi;Birim Süre (dk);Hazırlık Süresi (dk)\n"
                                "P1;10;Torna;CNC;6;30\nP1;20;Montaj;MON;3;0\n")
            m = yaz(t, "m.csv", "Makine;İş Merkezi;Vardiya;Verimlilik\nCNC-1;CNC;1;100\nCNC-2;CNC;1;100\nMON-1;MON;1;100\n")
            c = main.calistir(s, r, m, t / "u.xlsx", PZT)
        op = {(o["siparis"], o["op"]): o for o in c["operasyonlar"]}
        # A (termin önce): CNC-1 08:00–12:30 (30 + 40×6 = 270 dk), Montaj 12:30–14:30
        self.assertEqual((op[("A", "Torna")]["makine"], op[("A", "Torna")]["bitis"]), ("CNC-1", datetime(2026, 11, 2, 12, 30)))
        self.assertEqual(op[("A", "Montaj")]["bitis"], datetime(2026, 11, 2, 14, 30))
        # B: 630 dk; CNC-2 boşta → Pzt 480 + Salı 150 → 10:30 (CNC-1'de 15:00 olurdu); Montaj Salı 10:30–15:30
        self.assertEqual((op[("B", "Torna")]["makine"], op[("B", "Torna")]["bitis"]), ("CNC-2", datetime(2026, 11, 3, 10, 30)))
        self.assertEqual(op[("B", "Montaj")]["bitis"], datetime(2026, 11, 3, 15, 30))
        self.assertTrue(any("P2" in u for u in c["uyarilar"]))                 # rotası olmayan ürün
        self.assertEqual([x["gecikme"] for x in c["plan"]], [0, 0])

    def test_verimlilik_ve_gecikme(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            s = yaz(t, "s.csv", "Sipariş No;Ürün;Miktar;Termin\nA;P1;100;02.11.2026\n")
            r = yaz(t, "r.csv", "Ürün;İş Merkezi;Birim Süre (dk)\nP1;CNC;4\n")
            m = yaz(t, "m.csv", "Makine;İş Merkezi;Verimlilik\nCNC-1;CNC;%80\n")
            c = main.calistir(s, r, m, t / "u.xlsx", PZT)
        self.assertEqual(c["plan"][0]["bitis"], datetime(2026, 11, 3, 8, 20))    # 400 / 0,8 = 500 dk
        self.assertEqual(c["plan"][0]["gecikme"], 1)

    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            c = main.calistir(ORNEK / "siparisler.csv", ORNEK / "rotalar.csv", ORNEK / "makineler.csv", Path(tmp) / "u.xlsx", PZT)
        self.assertEqual(len(c["plan"]), 6)
        self.assertEqual(c["plan"][0]["no"], "İE-2616")                         # en erken termin
        self.assertEqual(max(c["doluluk"], key=c["doluluk"].get), "CNC-01")
        # Aynı makinede operasyonlar çakışmaz
        for mk in {o["makine"] for o in c["operasyonlar"]}:
            ops = sorted((o for o in c["operasyonlar"] if o["makine"] == mk), key=lambda o: o["baslangic"])
            for a, b in zip(ops, ops[1:]):
                self.assertLessEqual(a["bitis"], b["baslangic"])


if __name__ == "__main__":
    unittest.main()
