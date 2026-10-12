import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

ORNEK = main.ORNEK


def dn(*satirlar):
    f = lambda v: None if v is None else D(v)   # noqa: E731
    return [{"donem": f"D{i}", "pv": D(pv), "ev": f(ev), "ac": f(ac)} for i, (pv, ev, ac) in enumerate(satirlar, 1)]


PLAN = lambda: dn((1000, 900, 1000), (2000, 1800, 2000), (3000, 2400, 3000), (3000, None, None), (2000, None, None), (1000, None, None))   # noqa: E731


class AnalizTesti(unittest.TestCase):
    def test_gostergeler(self):
        r = main.analiz(PLAN())
        o = r["ozet"]
        self.assertEqual((o["donem"], o["bac"], o["pv"], o["ev"], o["ac"], o["sv"], o["cv"], o["spi"], o["cpi"], o["tamamlanma"], o["planlanan"]),
                         ("D3", D(12000), D(6000), D(5100), D(6000), D(-900), D(-900), D("0.85"), D("0.85"), D("42.50"), D("50.00")))
        self.assertEqual((o["eac"], o["etc"], o["vac"], o["eac_butce"], o["eac_ikili"], o["tcpi"], o["sure_tahmini"], o["program"], o["maliyet"], o["plan_donem"], o["gecen_donem"]),
                         (D("14117.65"), D("8117.65"), D("-2117.65"), D(12900), D("15550.17"), D("1.15"), D("7.06"), "Programın gerisinde", "Bütçenin üzerinde", 6, 3))
        self.assertEqual([(s["pv"], s["ev"], s["ac"], s.get("spi"), s.get("cpi"), s["ileri"]) for s in r["donemler"]][:4],
                         [(D(1000), D(900), D(1000), D("0.90"), D("0.90"), False), (D(3000), D(2700), D(3000), D("0.90"), D("0.90"), False), (D(6000), D(5100), D(6000), D("0.85"), D("0.85"), False),
                          (D(9000), None, None, None, None, True)])

    def test_kumulatif_bac_ve_iyi_durum(self):
        k = main.analiz(dn((1000, 900, 1000), (3000, 2700, 3000), (6000, 5100, 6000), (9000, None, None), (11000, None, None), (12000, None, None)), kumulatif=True)
        self.assertEqual(k["ozet"], main.analiz(PLAN())["ozet"])
        o = main.analiz(dn((1000, 1100, 1000), (1000, 1100, 900)))["ozet"]
        self.assertEqual((o["spi"], o["cpi"], o["eac"], o["vac"], o["program"], o["maliyet"], o["tcpi"]), (D("1.10"), D("1.16"), D("1727.27"), D("272.73"), "Programın ilerisinde", "Bütçenin altında", D("-2.00")))
        r = main.analiz(dn((1000, 500, 500)), D(4000))
        self.assertEqual((r["ozet"]["bac"], r["ozet"]["eac"], r["ozet"]["tamamlanma"], r["bac_farki"], r["ozet"]["program"], r["ozet"]["maliyet"]), (D(4000), D("4000.00"), D("12.50"), True, "Programın gerisinde", "Bütçede"))
        o = main.analiz(dn((1000, 0, 0)))["ozet"]
        self.assertEqual((o["spi"], o["cpi"], o["eac"], o["eac_ikili"], o["sure_tahmini"]), (D("0.00"), None, None, None, None))
        self.assertEqual(main.analiz(dn((1000, 900, 1200)))["ozet"]["tcpi"], None)                       # bütçe tükendi

    def test_hatalar(self):
        for d, kw in ((dn(), {}), (dn((1000, None, None)), {}), (dn((1000, None, None), (1000, 5, 5)), {}), (dn((1000, 900, 900), (500, 800, 800)), {"kumulatif": True}), (dn((0, 0, 0)), {})):
            with self.assertRaises(ValueError):
                main.analiz(d, **kw)


class DosyaTesti(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            cikti = Path(tmp) / "e.xlsx"
            r = main.calistir(ORNEK / "donemler.csv", cikti)
            o = r["ozet"]
            self.assertEqual((o["donem"], o["bac"], o["spi"], o["cpi"], o["eac"], o["vac"], o["tcpi"], r["uyarilar"]), ("2026-09", D(12000000), D("0.85"), D("0.85"), D("14117647.06"), D("-2117647.06"), D("1.15"), []))
            self.assertEqual(load_workbook(cikti).sheetnames, ["Özet", "Dönem Bazında", "Uyarılar"])

    def test_okuma_ve_cli(self):
        with tempfile.TemporaryDirectory() as tmp:
            y = Path(tmp) / "d.csv"
            y.write_text("Ay;PV;EV;AC\nOcak;1.000,00;900;1.000\nŞubat;1.000;950;\nMart;1.000;;\n", encoding="utf-8")
            l, uy = main.oku(y)
            self.assertEqual(([(d["donem"], d["pv"], d["ev"], d["ac"]) for d in l], [u["tur"] for u in uy]), ([("Ocak", D("1000.00"), D(900), D(1000)), ("Şubat", D(1000), D(950), D(0)), ("Mart", D(1000), None, None)], ["Dönem verisi eksik"]))
            c = str(Path(tmp) / "x.xlsx")
            arg = ["--donemler", str(y), "--cikti", c]
            self.assertEqual(sorted(u["tur"] for u in main.calistir(y, Path(c), D(2500))["uyarilar"]), ["BAC planlanan değer toplamından farklı", "Dönem verisi eksik"])
            self.assertEqual(main.main(arg + ["--bac", "3.500,00"]), 0)
            self.assertEqual(main.main(arg + ["--kumulatif"]), 1)                                   # dönemsel veriyi kümülatif sanınca PV azalır
            self.assertEqual(main.main(arg + ["--bac", "0"]), 2)
            self.assertEqual(main.main(["--donemler", str(Path(tmp) / "yok.csv")]), 1)
            for icerik in ("Dönem;PV;EV;AC\nOcak;çok;1;1\n", "Dönem;PV;EV;AC\nOcak;10;-1;1\n", "Dönem;PV\n"):
                y.write_text(icerik, encoding="utf-8")
                self.assertEqual(main.main(arg), 1)
            y.write_text("Dönem;PV;EV;AC\nOcak;100;120;150\n", encoding="utf-8")
            self.assertEqual(sorted(u["tur"] for u in main.calistir(y, Path(c))["uyarilar"]), ["Bütçe tükendi", "Kazanılmış değer bütçeyi aşıyor"])
            self.assertEqual(main.main(["--cikti", c]), 0)


if __name__ == "__main__":
    unittest.main()
