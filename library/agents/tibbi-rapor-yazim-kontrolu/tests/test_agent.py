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
import belge  # noqa: E402
import llm  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

V = KLASOR / "ornek_veri"
HASTA, RAPOR = V / "hasta.txt", V / "rapor.txt"
MESAJLAR = []


def b(tur, alinti, oneri, aciklama="", sor=False):
    return {"tur": tur, "alinti": alinti, "oneri": oneri, "aciklama": aciklama, "hekime_sor": sor}


def cevap(dogru=True) -> dict:
    t = {"bulgular": [b("Yazım", "Medial meniksüs", "Medial menisküs", "Terim yazımı"),
                      b("Taraf tutarlılığı", "Sol dizde", "Sağ dizde", "Rapor sağ diz; hekime sorulmalı", True),
                      b("Kısaltma", "yağ baskılı PD sekanslarda", "yağ baskılı proton dansite (PD) sekanslarda", "Kısaltma ilk geçtiği yerde açılmalı")]}
    if not dogru:
        t["bulgular"] += [b("Yazım", "medial mensküs", "medial menisküs"), b("Birim / ölçü", "8 mm çaplı", "18 mm çaplı"),
                          b("Yazım", "Minimal efüzyon", "Belirgin sıvı birikimi"), b("Yazım", "Medial meniksüs", "Medial menisküs")]
    return t


def sahte_json_iste(sistem, kullanici, sema, max_tokens=16000):
    MESAJLAR.append(kullanici)
    return cevap("<duzeltme>" in kullanici)


def veri():
    h = agent.hasta_oku(HASTA)
    m = belge.metin_oku(RAPOR)
    return h, m, agent.gizle(m, h)


class KodTesti(unittest.TestCase):
    def test_kod_kontrolleri(self):
        h, m, g = veri()
        self.assertEqual(agent.yas(date(1979, 2, 3), date(2026, 10, 15)), 47)
        k = agent.kod_kontrolleri(m, h)
        self.assertEqual([x["alinti"] for x in k], ["45 yaş", "kadın hasta", "12.11.2026", "2026-104587", "Sol dizde suprapatellar bursada minimal efüzyon mevcuttur."])
        self.assertTrue(all(x["hekime_sor"] for x in k))
        self.assertNotIn("10000000146", g)
        self.assertNotIn("2026-104587", g)
        h2 = dict(h, cinsiyet="Kadın", tur="Sol diz MR")
        k2 = agent.kod_kontrolleri("Prostat normal. Sağ dizde ödem. SAĞ DİZ.", h2)
        self.assertEqual([x["alinti"] for x in k2], ["Prostat", "Sağ dizde ödem.", "SAĞ DİZ."])

    def test_denetim(self):
        h, m, g = veri()
        self.assertEqual(agent.denetle(cevap(), g), [])
        s = agent.denetle(cevap(False), g)
        for parca in ("'medial mensküs'", "sayıyı değiştiriyor", "çok farklı", "iki kez"):
            self.assertTrue(any(parca in x for x in s), (parca, s))
        self.assertEqual(len(s), 4, s)


class AgentTesti(unittest.TestCase):
    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_dongusu(self, _):
        MESAJLAR.clear()
        with tempfile.TemporaryDirectory() as tmp:
            cikti = Path(tmp) / "r.xlsx"
            s = agent.calistir(HASTA, RAPOR, cikti, evet=True)
            self.assertEqual((len(MESAJLAR), s["kalan"], s["tur"]), (2, [], 1))
            for gizli in ("Ali Deneme", "10000000146", "2026-104578", "03.02.1979"):
                self.assertNotIn(gizli, "".join(MESAJLAR))
            self.assertEqual((len(s["bulgular"]), s["uygulanan"]), (7, 2))
            self.assertIn("Medial menisküs arka boynuzda", s["oneri"])
            self.assertIn("Sol dizde", s["oneri"])  # hekime sorulacak öneri uygulanmaz
            self.assertIn("proton dansite (PD)", cikti.with_suffix(".txt").read_text(encoding="utf-8"))
            ws = load_workbook(cikti)["Bulgular"]
            self.assertEqual((ws["B2"].value, ws["G2"].value, ws["B7"].value), ("Kod", "Evet", "Model"))

    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_kapali_ve_cli(self, _):
        with tempfile.TemporaryDirectory() as tmp:
            s = agent.calistir(HASTA, RAPOR, Path(tmp) / "r.xlsx", tur=0, evet=True)
            self.assertEqual(len(s["kalan"]), 4)
            self.assertEqual(agent.main(["--evet", "--cikti", str(Path(tmp) / "c.xlsx")]), 0)
            bozuk = Path(tmp) / "h.txt"
            bozuk.write_text("Hasta Adı: X\nCinsiyet: ?\n", encoding="utf-8")
            self.assertEqual(agent.main(["--evet", "--hasta", str(bozuk)]), 1)


if __name__ == "__main__":
    unittest.main()
