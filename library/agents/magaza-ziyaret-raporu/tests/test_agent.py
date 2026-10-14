"""Agent testleri gerçek API çağırmaz: llm.json_iste sahte bir fonksiyonla değiştirilir."""
import os
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))
os.environ.setdefault("ANTHROPIC_API_KEY", "test-anahtari")  # sahte; testler API çağırmaz

import agent  # noqa: E402
import llm  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

KONTROL = KLASOR / "ornek_veri" / "kontrol_listesi.csv"
NOTLAR = KLASOR / "ornek_veri" / "ziyaret_notlari.txt"
ZIYARET = date(2026, 10, 14)
MESAJLAR = []


def aksiyon(maddeler, metin, termin="21.10.2026", oncelik="Orta", sorumlu="Mağaza müdürü"):
    return {"maddeler": maddeler, "aksiyon": metin, "sorumlu": sorumlu, "termin": termin, "oncelik": oncelik}


def taslak(dogru: bool) -> dict:
    a = [aksiyon(["K2"], "Teknik servis kaydı açılıp sol vitrin spotları değiştirilecek"), aksiyon(["K4"], "Triko masası planograma göre yeniden düzenlenecek"),
         aksiyon(["K5"], "Trençkot M ve L bedenleri depodan reyona alınacak; günlük beden kontrolü yapılacak", oncelik="Yüksek"),
         aksiyon(["K8"], "Sevkiyat kolileri koridordan kaldırılıp depoya yerleştirilecek", "16.10.2026"),
         aksiyon(["K11"], "Eski fiyatlı etiketler değiştirilecek, 20 ürünlük etiket-kasa karşılaştırması yapılacak", "15.10.2026", "Yüksek"),
         aksiyon(["K15"], "Yaka kartları merkezden talep edilecek", sorumlu="Mağaza müdürü"),
         aksiyon(["K17"], "Acil çıkış önü için zemin işaretlemesi yapılacak, günlük açılış kontrolüne eklenecek", "16.10.2026", "Yüksek" if dogru else "Orta")]
    if dogru:
        a.append(aksiyon(["K10"], "İade kolileri için kargo randevusu alınıp koliler gönderilecek", "18.10.2026", "Yüksek", "Bölge lojistik"))
    else:
        a[1]["termin"] = "10.10.2026"
    return {"ozet": "Genel uyum %70,6; en zayıf bölüm Stok ve depo.", "guclu_yonler": [{"metin": "Karşılama standardı iyi", "maddeler": ["K14"] if dogru else ["K14", "K15"]}],
            "aksiyonlar": a, "gorusulenler": ["Cumartesi için 1 ek part-time personel talebi"], "takip_konulari": ["Beden dizilimi", "Acil çıkış önü"]}


def sahte_json_iste(sistem, kullanici, sema, max_tokens=16000):
    MESAJLAR.append(kullanici)
    return taslak("<duzeltme>" in kullanici)


class KodTesti(unittest.TestCase):
    def test_puanlama(self):
        m, h = agent.kontrol_oku(KONTROL)
        p = agent.puanla(m)
        self.assertEqual((len(m), h, m[16]["id"], m[16]["sonuc"]), (19, [], "K17", "Uygun değil"))
        # 18 değerlendirilen madde: 10 uygun + 4 × 0,5 = 12 → %66,7
        self.assertEqual((p[-1]["uygun"], p[-1]["kismen"], p[-1]["degil"], p[-1]["puan"]), (10, 4, 4, 66.7))
        self.assertEqual([(x["bolum"], x["puan"]) for x in p[:2]], [("Vitrin ve giriş", 83.3), ("Satış alanı", 62.5)])
        self.assertEqual([agent.sonuc(x) for x in ("Uygun", "kısmi", "Uygun Değil", "DD", "2", "x")], ["Uygun", "Kısmen", "Uygun değil", "DD", "Uygun", None])
        self.assertEqual(agent.ziyaret_tarihi(NOTLAR.read_text(encoding="utf-8")), ZIYARET)

    def test_denetim(self):
        m, _ = agent.kontrol_oku(KONTROL)
        self.assertEqual(agent.denetle(taslak(True), m, ZIYARET), [])
        s = agent.denetle(taslak(False), m, ZIYARET)
        self.assertEqual(len(s), 4)
        self.assertTrue(any("K10" in x and "aksiyon yok" in x for x in s) and any("K17" in x and "Yüksek" in x for x in s)
                        and any("10.10.2026" in x for x in s) and any("K15 'Kısmen'" in x for x in s))
        t = taslak(True)
        t["aksiyonlar"][0].update({"termin": "x", "sorumlu": " ", "maddeler": ["K2", "K99"]})
        self.assertEqual(len(agent.denetle(t, m, ZIYARET)), 3)
        self.assertEqual(len(agent.denetle(taslak(True), m, ZIYARET, azami=3)), 5)


class AgentTesti(unittest.TestCase):
    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_dongusu(self, _):
        MESAJLAR.clear()
        with tempfile.TemporaryDirectory() as tmp:
            cikti = Path(tmp) / "m.xlsx"
            s = agent.calistir(KONTROL, NOTLAR, cikti, evet=True)
            self.assertEqual((len(MESAJLAR), s["kalan"], s["tur"]), (2, [], 1))
            self.assertIn("K17 | Güvenlik ve İSG | Acil çıkış", MESAJLAR[0])
            self.assertIn("K10", MESAJLAR[1].split("<duzeltme>")[1])
            wb = load_workbook(cikti)
            self.assertEqual(wb.sheetnames, ["Rapor", "Aksiyon Planı", "Bölüm Puanları", "Kontrol Listesi"])
            self.assertEqual(wb["Aksiyon Planı"]["A2"].value, "Yüksek")
            self.assertIn("## Aksiyon planı", s["md"].read_text(encoding="utf-8"))

    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_kapali_ve_cli(self, _):
        with tempfile.TemporaryDirectory() as tmp:
            s = agent.calistir(KONTROL, NOTLAR, Path(tmp) / "m.xlsx", tur=0, evet=True)
            self.assertEqual(len(s["kalan"]), 4)
            self.assertEqual(agent.main(["--evet", "--cikti", str(Path(tmp) / "c.xlsx")]), 0)
            self.assertEqual(agent.main(["--tarih", "x"]), 2)
            bos = Path(tmp) / "b.csv"
            bos.write_text("Bölüm;Madde\nA;B\n", encoding="utf-8")
            self.assertEqual(agent.main(["--kontrol", str(bos), "--evet"]), 1)
            kopya = Path(tmp) / "k.csv"
            kopya.write_text(KONTROL.read_text(encoding="utf-8"), encoding="utf-8")
            self.assertEqual(agent.main(["--kontrol", str(kopya), "--evet"]), 1)
            self.assertEqual(agent.main(["--kontrol", str(kopya), "--tarih", "14.10.2026", "--evet", "--cikti", str(Path(tmp) / "d.xlsx")]), 0)


if __name__ == "__main__":
    unittest.main()
