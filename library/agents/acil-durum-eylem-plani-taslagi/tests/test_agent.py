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

GIRDI = KLASOR / "ornek_veri" / "isyeri.txt"
MESAJLAR = []


def m(ekip, adim):
    return {"ekip": ekip, "adim": adim}


def taslak(dogru=True) -> dict:
    t = {"senaryolar": [
        {"ad": "Üretim holü ve depo yangını", "tur": "Yangın", "dayanak": ["R1", "R2", "B1", "B2", "B3", "E1", "E2"],
         "onleyici_tedbirler": ["Solvent kimyasal dolabında kapalı tutulur.", "Granül stokları kaçış yollarını kapatmayacak şekilde istiflenir."],
         "mudahale": [m("Tüm çalışanlar", "Yangını gören alarm butonuna basar ve 112'yi arar."), m("Söndürme", "KKT yangın söndürücülerle ilk müdahaleyi yapar."),
                      m("Koruma", "Elektrik ana kesme şalterini kapatır ve itfaiyeyi karşılar.")]},
        {"ad": "LPG kaçağı ve kompresör tankı patlaması", "tur": "Patlama", "dayanak": ["R3", "R6", "B4", "B5", "E6"],
         "onleyici_tedbirler": ["Mutfakta gaz algılama dedektörü sağlanmalı.", "Hava tankının periyodik kontrolleri takip edilir."],
         "mudahale": [m("Koruma", "LPG ana vanasını kapatır ve bölgeyi boşaltır."), m("İlk yardım", "Yaralılara ilk yardım uygular.")]},
        {"ad": "Deprem", "tur": "Doğal afet", "dayanak": ["R4", "B2"],
         "onleyici_tedbirler": ["Raf sistemleri sabitlenir."],
         "mudahale": [m("Tüm çalışanlar", "Sarsıntı sırasında çök-kapan-tutun yapar."), m("Kurtarma", "Sarsıntı sonrası arama ve kurtarma yapar.")]},
        {"ad": "Makinede sıkışma ve yanık", "tur": "Diğer", "dayanak": ["R5", "B1"],
         "onleyici_tedbirler": ["Makine koruyucuları devrede tutulur."],
         "mudahale": [m("İlk yardım", "Yanığı soğutur ve 112'yi arar.")]}],
        "tahliye": ["Sesli alarmla tahliye başlatılır.", "Gebe ve tekerlekli sandalye kullanan çalışanlara refakat edilir.",
                    "Ziyaretçi ve tedarikçi şoförleri koruma ekibince yönlendirilir.", "Herkes toplanma yerine gider.",
                    "Koruma ekibi toplanma yerinde sayım yapar; eksik kişi kurtarma ekibine bildirilir."]}
    if not dogru:
        t["senaryolar"] = t["senaryolar"][:1] + t["senaryolar"][2:3]
        t["senaryolar"][0]["mudahale"][1] = m("Tüm çalışanlar", "5 kişilik söndürme ekibi 110'u arar.")
        t["tahliye"] = ["Herkes toplanma yerine gider."]
    return t


def sahte_json_iste(sistem, kullanici, sema, max_tokens=16000):
    MESAJLAR.append(kullanici)
    return taslak("<duzeltme>" in kullanici)


class KodTesti(unittest.TestCase):
    def test_hesap(self):
        n = agent.isyeri_oku(GIRDI)
        self.assertEqual((n["vardiya"], n["toplam"], len(n["riskler"]), len(n["ekipman"])), ([("Gündüz", 85), ("Gece", 32)], 117, 6, 7))
        e = agent.ekip_hesapla(n)
        self.assertEqual(e["vardiyalar"][0]["ekipler"], {"Söndürme": 3, "Kurtarma": 3, "Koruma": 3, "İlk yardım": 6})
        self.assertEqual(e["vardiyalar"][1]["ekipler"], {"Söndürme": 1, "Kurtarma": 1, "Koruma": 1, "İlk yardım": 3})
        self.assertEqual((e["gecerlilik"], e["tatbikat"]), (date(2030, 10, 15), date(2027, 10, 15)))
        self.assertEqual(agent.dokuman_eksikleri(n), [])
        n["sinif"], n["vardiya"], n["toplam"] = "Çok tehlikeli", [("Tek", 8)], 8
        self.assertEqual(agent.ekip_hesapla(n)["vardiyalar"][0]["ekipler"], {"Söndürme + Kurtarma + Koruma (tek kişi)": 1, "İlk yardım": 1})
        self.assertEqual(agent.ekip_hesapla(n)["gecerlilik"], date(2028, 10, 15))

    def test_denetim(self):
        n = agent.isyeri_oku(GIRDI)
        self.assertEqual(agent.denetle(taslak(), n), [])
        s = agent.denetle(taslak(False), n)
        for parca in ("söndürme ekibine", "110", "5 (ekip", "R3", "R5", "R6", "sayım", "refakat", "ziyaretçi"):
            self.assertTrue(any(parca in x for x in s), (parca, s))
        self.assertEqual(len(s), 9, s)


class AgentTesti(unittest.TestCase):
    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_dongusu(self, _):
        MESAJLAR.clear()
        with tempfile.TemporaryDirectory() as tmp:
            cikti = Path(tmp) / "p.xlsx"
            s = agent.calistir(GIRDI, cikti, evet=True)
            self.assertEqual((len(MESAJLAR), s["kalan"], s["tur"]), (2, [], 1))
            for gizli in ("Örnek Plastik", "Organize Sanayi", "Fabrika Müdürü"):
                self.assertNotIn(gizli, "".join(MESAJLAR))
            metin = cikti.with_suffix(".txt").read_text(encoding="utf-8")
            self.assertIn("Geçerlilik tarihi: 15.10.2030", metin)
            self.assertIn("Gündüz (85 çalışan): Söndürme 3, Kurtarma 3, Koruma 3, İlk yardım 6", metin)
            self.assertIn("İlk tatbikat en geç 15.10.2027", metin)
            ws = load_workbook(cikti)["Ekipler"]
            self.assertEqual((ws["C2"].value, ws["D2"].value), ("Söndürme", 3))

    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_kapali_ve_cli(self, _):
        with tempfile.TemporaryDirectory() as tmp:
            s = agent.calistir(GIRDI, Path(tmp) / "p.xlsx", tur=0, evet=True)
            self.assertEqual(len(s["kalan"]), 9)
            self.assertEqual(agent.main(["--evet", "--cikti", str(Path(tmp) / "c.xlsx")]), 0)
            bozuk = Path(tmp) / "b.txt"
            bozuk.write_text("Unvan: X\nTehlike Sınıfı: Orta\nRiskler:\n- yangın\n", encoding="utf-8")
            self.assertEqual(agent.main(["--evet", "--girdi", str(bozuk)]), 1)


if __name__ == "__main__":
    unittest.main()
