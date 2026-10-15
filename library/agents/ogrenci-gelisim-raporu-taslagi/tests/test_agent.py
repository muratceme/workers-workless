"""Agent testleri gerçek API çağırmaz: llm.json_iste sahte bir fonksiyonla değiştirilir."""
import os
import re
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))
os.environ.setdefault("ANTHROPIC_API_KEY", "test-anahtari")  # sahte; testler API çağırmaz

import agent  # noqa: E402
import llm  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

V = KLASOR / "ornek_veri"
NOTLAR, DEV, GOZ = V / "notlar.csv", V / "devamsizlik.csv", V / "gozlemler.txt"
MESAJLAR = []


def m(metin, *d):
    return {"metin": metin, "dayanak": list(d)}


def taslak(k, dogru=True) -> dict:
    if k == "Ö1":
        t = {"genel_degerlendirme": "{OGRENCI} bu dönem 76,2 ortalamayla dil derslerinde başarılı; matematikte desteğe ihtiyaç duyuyor.",
             "guclu_yonler": [m("İngilizce dersinde 91 puanla başarılı.", "İngilizce"), m("Sözlü anlatımda kendini iyi ifade ediyor.", "G1", "Türk Dili ve Edebiyatı"),
                              m("Grup çalışmalarında uyumlu.", "G3")],
             "gelisime_acik": [m("Matematik puanı 71'den 58'e düştü (13 puan); ödevlerini eksik getiriyor.", "Matematik", "G2"), m("Fizik puanı 62.", "Fizik")],
             "oneriler": [m("Matematik için haftalık çalışma planı ve etüt desteği.", "Matematik", "G2")],
             "veli_sorulari": ["Evde matematik çalışmasına ne kadar zaman ayırıyor?"]}
        if not dogru:
            t["guclu_yonler"].append(m("Fizikte istikrarlı.", "Fizik"))
            t["gelisime_acik"] = t["gelisime_acik"][1:]
            t["genel_degerlendirme"] = "{OGRENCI} 79 ortalamaya sahip; derste dikkat eksikliği olabilir."
        return t
    return {"genel_degerlendirme": "{OGRENCI} sayısal derslerde güçlü; tarih dersinde desteğe ihtiyaç duyuyor.",
            "guclu_yonler": [m("Matematik dersinde 92 puan; problem çözmeye istekli.", "Matematik", "G1")],
            "gelisime_acik": [m("Tarih puanı 60; not tutmakta zorlanıyor, dikkati çabuk dağılıyor.", "Tarih", "G2")],
            "oneriler": [m("Tarih derslerinde not alma tekniklerine yönelik rehberlik.", "Tarih", "G2")], "veli_sorulari": ["Evde düzenli bir çalışma ortamı var mı?"]}


def sahte_json_iste(sistem, kullanici, sema, max_tokens=16000):
    MESAJLAR.append(kullanici)
    return taslak(re.search(r'kimlik="(Ö\d+)"', kullanici).group(1), "<duzeltme>" in kullanici)


def ogr():
    return agent.ogrencileri_oku(NOTLAR, DEV, GOZ, 10)


class KodTesti(unittest.TestCase):
    def test_hesap(self):
        o = ogr()
        self.assertEqual([(x["kimlik"], round(x["ortalama"], 1), x["dusenler"]) for x in o], [("Ö1", 76.2, ["Matematik"]), ("Ö2", 75.0, [])])
        self.assertEqual((o[0]["ozursuz"], len(o[0]["gozlemler"])), (3.0, 3))
        self.assertTrue(o[0]["gozlemler"][2].startswith("{OGRENCI} grup"))
        self.assertIn("- Matematik: önceki 71 → bu dönem 58 (değişim −13)", agent.ogrenci_metni(o[0]))

    def test_denetim(self):
        o = ogr()
        self.assertEqual(agent.denetle(o, {"Ö1": taslak("Ö1"), "Ö2": taslak("Ö2")}), [])
        s = agent.denetle(o, {"Ö1": taslak("Ö1", False), "Ö2": taslak("Ö2")})
        self.assertEqual(len(s), 4, s)
        self.assertTrue(all(x.startswith("Ö1:") for x in s) and any("Fizik" in x and "ortalamanın" in x for x in s) and any("Matematik" in x and "düşüş" in x for x in s)
                        and any("79" in x for x in s) and any("dikkat eksikli" in x for x in s))
        k = taslak("Ö2")
        k["gelisime_acik"][0]["metin"] = "Tarihte sınıfın en düşük puanı."
        k["oneriler"][0]["dayanak"] = ["Kimya", "G9"]
        s = agent.denetle(o, {"Ö1": taslak("Ö1"), "Ö2": k})
        self.assertTrue(any("kıyas" in x for x in s) and any("'Kimya'" in x for x in s) and any("G9" in x for x in s))


class AgentTesti(unittest.TestCase):
    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_dongusu(self, _):
        MESAJLAR.clear()
        with tempfile.TemporaryDirectory() as tmp:
            cikti = Path(tmp) / "g.xlsx"
            s = agent.calistir(NOTLAR, cikti, DEV, GOZ, evet=True)
            self.assertEqual((len(MESAJLAR), s["kalan"], s["tur"]), (3, [], 1))
            for gizli in ("Deniz", "Aydın", "Mert", "1021", "1034"):
                self.assertNotIn(gizli, "".join(MESAJLAR))
            self.assertIn('kimlik="Ö1"', MESAJLAR[2])
            r = s["dosyalar"][0].read_text(encoding="utf-8")
            self.assertTrue(r.startswith("ÖĞRENCİ GELİŞİM RAPORU (TASLAK) — Deniz Aydın · 9-A · No: 1021"))
            self.assertIn("Deniz Aydın bu dönem 76,2 ortalamayla", r)
            self.assertEqual(load_workbook(cikti)["Özet"]["E2"].value, "Matematik")

    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_kapali_ve_cli(self, _):
        with tempfile.TemporaryDirectory() as tmp:
            s = agent.calistir(NOTLAR, Path(tmp) / "g.xlsx", DEV, GOZ, tur=0, evet=True)
            self.assertEqual((len(s["kalan"]), len(s["dosyalar"])), (4, 1))
            self.assertEqual(agent.main(["--evet", "--cikti", str(Path(tmp) / "c.xlsx")]), 0)
            self.assertEqual(agent.main(["--evet", "--gozlemler", str(GOZ.with_name("yok.txt"))]), 1)


if __name__ == "__main__":
    unittest.main()
