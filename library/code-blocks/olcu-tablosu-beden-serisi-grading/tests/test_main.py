import sys
import tempfile
import unittest
from decimal import Decimal as D
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

ORNEK = KLASOR / "ornek_veri" / "erkek_tisort_kurallar.csv"
BEDEN = ["XS", "S", "M", "L", "XL", "XXL"]


class SeriTesti(unittest.TestCase):
    def test_yukari_asagi(self):
        k = {"ana": D(52), "artis": {"XS-S": D(2), "S-M": D(2), "M-L": D(2), "L-XL": D(3), "XL-XXL": D(3)}}
        self.assertEqual(main.seri_uret(k, BEDEN, "M"), {"XS": 48, "S": 50, "M": 52, "L": 54, "XL": 57, "XXL": 60})

    def test_ana_beden_uc(self):
        k = {"ana": D(40), "artis": {"36-38": D(1), "38-40": D(1)}}
        self.assertEqual(main.seri_uret(k, ["36", "38", "40"], "40"), {"36": 38, "38": 39, "40": 40})

    def test_kontrol(self):
        k = {"tol": D(1), "artis": {"S-M": D("0.5"), "M-L": D(-1)}}
        notlar = main.kontrol(k, {"S": D(10), "M": D("10.5"), "L": D("9.5")}, ["S", "M", "L"])
        self.assertIn("büyük beden küçükten küçük (negatif artış)", notlar)
        self.assertIn("artış toleranstan küçük: komşu bedenler ölçüyle ayırt edilemeyebilir", notlar)

    def test_yuvarlama(self):
        self.assertEqual(main.yuvarla(D("47.26"), D("0.5")), D("47.5"))
        self.assertEqual(main.yuvarla(D("47.24"), D("0.5")), D("47.0"))


class UctanUcaTest(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(ORNEK, Path(tmp) / "o.xlsx", BEDEN, "M", inc=True)
        t = {x["kural"]["kod"]: x for x in s["tablo"]}
        self.assertEqual([float(v) for v in t["A"]["seri"].values()], [48, 50, 52, 54, 57, 60])
        self.assertEqual([float(v) for v in t["D"]["seri"].values()], [44, 45, 46, 47.5, 49, 50.5])
        self.assertEqual([float(v) for v in t["G"]["seri"].values()], [17, 17.5, 18, 18.5, 19, 19.5])   # tek 'Artış' sütunu
        self.assertEqual([x for x in s["tablo"] if x["notlar"]], [])
        self.assertEqual(s["uyarilar"], [])

    def test_eksik_gecis_ve_hatali_ana(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "k.csv").write_text("Ölçü Noktası;Ana Beden;S-M\nBel;40;2\n", encoding="utf-8")
            s = main.calistir(t / "k.csv", t / "o.xlsx", ["S", "M", "L"], "M")
            self.assertTrue(any("M-L" in u for u in s["uyarilar"]))
            with self.assertRaises(SystemExit):
                main.calistir(t / "k.csv", t / "o.xlsx", ["S", "M", "L"], "XL")


if __name__ == "__main__":
    unittest.main()
