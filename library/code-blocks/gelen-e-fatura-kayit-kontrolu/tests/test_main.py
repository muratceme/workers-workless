import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class UctanUcaTest(unittest.TestCase):
    def test_senaryolar(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK / "gelen_efaturalar", ORNEK / "alis_kayitlari.csv", Path(tmp) / "k.xlsx")
        self.assertEqual([e.no for e in s["defterde_yok"]], ["EAR2026000000031"])          # hiç işlenmemiş
        self.assertEqual(sorted(k.no for k in s["portalda_yok"]), ["A-001234", "ORN2026000000101"])  # kâğıt + mükerrer 2. kayıt
        farkli = {m.e.no: m.tutar_farki for m in s["tutar_farkli"]}
        self.assertEqual(farkli, {"ORN2026000000115": D("90.00")})
        yontem = {m.e.no: m.yontem for m in s["eslesen"]}
        self.assertTrue(yontem["ORN2026000000110"].startswith("VKN"))                       # no hatalı, tutar tuttu
        self.assertEqual(len(s["mukerrer"]), 2)
        ihr = next(m for m in s["eslesen"] if m.e.no == "IHR2026000000007")
        self.assertEqual(ihr.e.tutar, D("260837.50"))                                      # USD × kur

    def test_portal_excel_listesi(self):
        with tempfile.TemporaryDirectory() as tmp:
            liste = Path(tmp) / "portal.csv"
            liste.write_text("Fatura No;Fatura Tarihi;Gönderici VKN/TCKN;Gönderici Unvan;Vergiler Dahil Tutar\n"
                             "ORN2026000000102;05.09.2026;5260181599;Örnek Gıda;24.000,00\n"
                             "XYZ2026000000999;06.09.2026;5260181599;Örnek Gıda;1.000,00\n", encoding="utf-8")
            s = main.calistir(liste, ORNEK / "alis_kayitlari.csv", Path(tmp) / "k.xlsx")
        self.assertEqual([e.no for e in s["defterde_yok"]], ["XYZ2026000000999"])


if __name__ == "__main__":
    unittest.main()
