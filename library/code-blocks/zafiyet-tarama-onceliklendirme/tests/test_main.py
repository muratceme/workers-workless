import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri"


class KuralTesti(unittest.TestCase):
    def test_oncelik_tablosu(self):
        o = main.oncelik
        self.assertEqual(o(9.8, True, True, 0.9, 5, True)[0], "P1")         # KEV + internete açık
        self.assertEqual(o(5.0, True, False, 0.0, 2, False)[0], "P2")       # KEV ama önemsiz varlık
        self.assertEqual(o(9.1, False, False, 0.71, 5, False)[0], "P1")     # yüksek EPSS + kritik varlık
        self.assertEqual(o(9.8, False, False, 0.0, 5, False)[0], "P2")      # CVSS ≥ 9 + kritik
        self.assertEqual(o(9.8, False, False, 0.0, 2, False)[0], "P3")
        self.assertEqual(o(7.5, False, False, 0.0, 2, False)[0], "P4")
        self.assertEqual(o(5.3, False, False, 0.0, 3, False)[0], "P5")

    def test_tarih_bicimleri(self):
        for t in ("2026-10-08", "2026-10-08T10:00:00Z", "08.10.2026", "Oct 8, 2026 10:00:00 UTC"):
            self.assertEqual(main.tarih(t), date(2026, 10, 8))


class UctanUcaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with tempfile.TemporaryDirectory() as t:
            cls.s = main.calistir(ORNEK / "tarama.csv", Path(t) / "a.xlsx", ORNEK / "varliklar.csv", ORNEK / "kev_ornek.csv",
                                  ORNEK / "epss_ornek.csv", bugun=date(2026, 10, 8))
        cls.b = {(f["host"], f["cve"][0] if f["cve"] else f["ad"]): f for f in cls.s["bulgular"]}

    def test_bilgi_bulgulari_atlanir(self):
        self.assertEqual(len(self.s["bulgular"]), 12)

    def test_kev_epss_varlik(self):
        vpn = self.b[("10.0.1.10", "CVE-2099-1001")]
        self.assertEqual((vpn["oncelik"], vpn["kev"], vpn["fidye"], vpn["internet"]), ("P1", True, True, True))
        self.assertEqual(self.b[("10.0.2.21", "CVE-2099-1002")]["oncelik"], "P2")    # KEV ama kritiklik 3, kapalı ağ
        arama = self.b[("10.0.4.70", "CVE-2099-1009")]
        self.assertAlmostEqual(arama["epss"], 0.55)                                    # iki CVE'nin büyük EPSS'i

    def test_sla_ve_gecikme(self):
        vpn = self.b[("10.0.1.10", "CVE-2099-1001")]
        self.assertEqual(vpn["hedef"], date(2026, 10, 2))                              # 25.09 + 7 gün
        self.assertEqual(vpn["gecikme"], 6)
        self.assertEqual(vpn["kev_son"], date(2026, 10, 19))

    def test_duzeltme_paketi(self):
        p = self.s["paketler"][0]
        self.assertEqual((p["oncelik"], len(p["hostlar"]), p["internet"]), ("P1", 2, 2))
        self.assertEqual(len(self.s["paketler"]), 8)

    def test_kaynaksiz_calisma(self):
        with tempfile.TemporaryDirectory() as t:
            s = main.calistir(ORNEK / "tarama.csv", Path(t) / "a.xlsx", bugun=date(2026, 10, 8))
        self.assertEqual(len(s["uyarilar"]), 3)
        self.assertFalse(any(f["kev"] for f in s["bulgular"]))


if __name__ == "__main__":
    unittest.main()
