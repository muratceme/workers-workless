import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal as D
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import main  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

ORNEK = main.ORNEK
BUGUN = date(2026, 10, 14)


def a(plaka, son, son_km=None, onceki=None, onceki_km=None, akm=None, aay=None, izin=True, tel="0532"):
    return {"plaka": plaka, "musteri": "Ali Veli", "telefon": tel, "model": "", "son_tarih": date(*son), "onceki_tarih": date(*onceki) if onceki else None,
            "son_km": None if son_km is None else D(son_km), "onceki_km": None if onceki_km is None else D(onceki_km), "aralik_km": None if akm is None else D(akm),
            "aralik_ay": None if aay is None else D(aay), "izin": izin}


class HesapTesti(unittest.TestCase):
    def test_vadeler(self):
        l = [a("KM", (2026, 6, 1), 30000, (2025, 6, 1), 6000, 15000, 12),           # 24.000 km / 365 gün = 65,8 km/gün → 15.000 km 228 günde: 01.06.2026 + 228 = 15.01.2027
             a("SURE", (2025, 11, 1), 20000, (2024, 11, 1), 26000 - 6000, 15000, 12),  # 6.000 km/yıl → süre önce: 01.11.2026
             a("GEC", (2025, 1, 10), 10000, akm=15000, aay=12, izin=False),           # varsayılan 40 km/gün → 375 gün → 20.01.2026; süre 10.01.2026
             a("YOK", (2026, 9, 1), 5000, tel=""), a("ESKI", (2024, 1, 1), aay=12)]
        r = main.hesapla(l, BUGUN, 30, mesaj="{plaka}: {tarih}")
        x = {v["plaka"]: v for v in r["araclar"]}
        self.assertEqual((x["KM"]["gunluk_km"], x["KM"]["km_kaynak"], x["KM"]["vade"], x["KM"]["neden"], x["KM"]["tahmini_km"], x["KM"]["hedef_km"], x["KM"]["durum"]),
                         (D("65.8"), "Geçmiş bakım", date(2027, 1, 15), "Km", 38883, 45000, "İleride"))
        self.assertEqual((x["SURE"]["vade"], x["SURE"]["neden"], x["SURE"]["kalan"], x["SURE"]["durum"]), (date(2026, 11, 1), "Süre", 18, "Yaklaşıyor"))
        self.assertEqual((x["GEC"]["vade"], x["GEC"]["kalan"], x["GEC"]["durum"]), (date(2026, 1, 10), -277, "Gecikmiş"))
        self.assertEqual((x["YOK"]["vade"], x["YOK"]["neden"]), (date(2027, 9, 1), "Süre"))
        self.assertEqual((x["ESKI"]["vade"], x["ESKI"]["tahmini_km"], x["ESKI"]["durum"]), (date(2025, 1, 1), None, "Gecikmiş"))
        self.assertEqual([v["plaka"] for v in r["liste"]], ["ESKI", "GEC", "SURE"])
        self.assertEqual(r["mesajlar"], [{"plaka": "ESKI", "telefon": "0532", "mesaj": "ESKI: geldi"}, {"plaka": "SURE", "telefon": "0532", "mesaj": "SURE: 01.11.2026 civarında geliyor"}])
        self.assertEqual(r["ozet"], {"Gecikmiş": 2, "Yaklaşıyor": 1, "İleride": 2, "Hesaplanamadı": 0})
        self.assertEqual(sorted((u["tur"], u["kim"]) for u in r["uyarilar"]),
                         [("Uzun süredir gelmiyor", "ESKI"), ("Uzun süredir gelmiyor", "GEC"), ("Varsayılan bakım aralığı", "2 araç"), ("İleti izni yok", "1 araç")])
        y = main.hesapla([a("T", (2026, 1, 1), 100, (2026, 2, 1), 50, 10000, 12)], BUGUN)
        self.assertEqual(y["uyarilar"][0]["tur"], "Km / tarih tutarsız")
        self.assertEqual(main.hesapla([a("N", (2026, 1, 1))], BUGUN, aralik_km=None, aralik_ay=None)["araclar"][0]["durum"], "Hesaplanamadı")
        with self.assertRaises(ValueError):
            main.hesapla([], BUGUN)

    def test_ay_ekle(self):
        self.assertEqual([main.ay_ekle(date(2026, 1, 31), 1), main.ay_ekle(date(2025, 11, 15), 12), main.ay_ekle(date(2026, 12, 31), 2), main.ay_ekle(date(2024, 2, 29), 12)],
                         [date(2026, 2, 28), date(2026, 11, 15), date(2027, 2, 28), date(2025, 2, 28)])


class DosyaTesti(unittest.TestCase):
    def test_ornek(self):
        with tempfile.TemporaryDirectory() as tmp:
            cikti = Path(tmp) / "t.xlsx"
            r = main.calistir(ORNEK / "araclar.csv", cikti, BUGUN, firma="Servis")
            self.assertEqual((len(r["araclar"]), r["ozet"]["Gecikmiş"], r["ozet"]["Yaklaşıyor"], len(r["mesajlar"])), (91, 22, 10, 23))
            self.assertTrue(all(m["mesaj"].endswith("Servis") for m in r["mesajlar"]))
            self.assertEqual(load_workbook(cikti).sheetnames, ["Hatırlatma Listesi", "Tüm Araçlar", "Mesajlar", "Uyarılar"])
            self.assertEqual(len((Path(tmp) / "t_mesajlar.csv").read_text(encoding="utf-8-sig").splitlines()), 24)

    def test_okuma_ve_cli(self):
        with tempfile.TemporaryDirectory() as tmp:
            y = Path(tmp) / "a.csv"
            y.write_text("Plaka;GSM;Son Servis Tarihi;Son Km;Bakım Km;İYS İzni\n34 ab 1;0532;01.01.2026;10.000;;Evet\n34 AB  1;;01.01.2026;;;\n34 C 2;;x;;;\n34 D 3;;01.01.2026;-5;;\n", encoding="utf-8")
            l, uy = main.oku(y)
            self.assertEqual(([(x["plaka"], x["son_km"], x["izin"]) for x in l], [u["tur"] for u in uy]), ([("34 AB 1", D(10000), True)], ["Plaka iki kez yazılmış", "Satır okunamadı", "Satır okunamadı"]))
            c = str(Path(tmp) / "x.xlsx")
            arg = ["--araclar", str(y), "--cikti", c]
            self.assertEqual(main.main(arg + ["--tarih", "14.10.2026", "--pencere", "60", "--gunluk-km", "50", "--aralik-km", "10000", "--aralik-ay", "6", "--mesaj", "{plaka} {tarih}"]), 0)
            self.assertEqual(main.main(arg), 0)
            self.assertEqual([main.main(arg + v) for v in (["--tarih", "x"], ["--gunluk-km", "0"], ["--aralik-ay", "0"], ["--pencere", "-1"], ["--mesaj", "{ad}"])], [2] * 5)
            self.assertEqual(main.main(["--araclar", str(Path(tmp) / "yok.csv")]), 1)
            y.write_text("Plaka;Son Servis Tarihi\nA;x\n", encoding="utf-8")
            self.assertEqual(main.main(arg), 1)
            self.assertEqual(main.main(["--cikti", c]), 0)


if __name__ == "__main__":
    unittest.main()
