"""Agent testleri gerçek API çağırmaz: llm.json_iste sahte bir fonksiyonla değiştirilir."""
import os
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
HASTA, NOT, TET = V / "hasta.txt", V / "notlar.txt", V / "tetkikler.csv"
MESAJLAR = []


def b(metin, *d):
    return {"metin": metin, "dayanak": list(d)}


def taslak(dogru=True) -> dict:
    t = {"oyku": b("Hasta 1 gündür olan karın ağrısı, bulantı ve iştahsızlık şikâyetiyle başvurdu; ağrı göbek çevresinden sağ alt kadrana yerleşmişti.", "N1"),
         "fizik_muayene": b("Sağ alt kadranda hassasiyet, defans ve rebound mevcuttu; ateş 38,2 °C idi.", "N2"),
         "yapilan_islemler": b("Batın USG'de apendiks çapı 9 mm ölçüldü. 11.10.2026 tarihinde laparoskopik apendektomi yapıldı; seftriakson 1 g IV 1x1 "
                               "ve metronidazol 500 mg IV 3x1 verildi.", "N3", "N4", "N5"),
         "klinik_seyir": b("Başvuruda lökosit 16,8, nötrofil 13,9 10^3/µL ve CRP 48 mg/L yüksekti. Postoperatif 1. gün ateşi olmadı, oral beslenmeye "
                           "başlandı. 13.10.2026'da lökosit 9,1'e, CRP 21 mg/L'ye geriledi.", "T1", "T2", "T3", "T6", "T7", "N5", "N6"),
         "cikis_durumu": b("Genel durumu iyi olarak 14.10.2026 tarihinde taburcu edildi.", "N7"),
         "oneriler": b("7. gün dikiş kontrolü için genel cerrahi polikliniğine başvurması; ağır kaldırmaması, ateş ve yara yerinde kızarıklıkta acile "
                       "başvurması önerildi.", "N7")}
    if not dogru:
        t["klinik_seyir"] = b("Postoperatif seyir sorunsuz iyileşme şeklinde oldu; hasta tamamen iyileşti. 15.10.2026'da kontrol hemogram normaldi.", "N6", "T9")
        t["yapilan_islemler"]["metin"] += " K37 tanısıyla seftriakson 1,5 g verildi."
    return t


def sahte_json_iste(sistem, kullanici, sema, max_tokens=16000):
    MESAJLAR.append(kullanici)
    return taslak("<duzeltme>" in kullanici)


def veri():
    return agent.hasta_oku(HASTA), agent.notlari_oku(NOT), agent.tetkik_oku(TET)


class KodTesti(unittest.TestCase):
    def test_okuma(self):
        h, n, t = veri()
        self.assertEqual(((h["cikis_v"] - h["yatis_v"]).days, len(n["notlar"]), len(n["ilaclar"]), n["tanilar"][0][:5]), (3, 7, 2, "K35.8"))
        self.assertEqual([x["durum"] for x in t], ["Yüksek", "Yüksek", "Yüksek", "Normal", "", "Normal", "Yüksek"])
        self.assertEqual((agent.ref_coz("4-10"), agent.ref_coz("<5"), agent.ref_coz("> 60")), ((4.0, 10.0), (None, 5.0), (60.0, None)))

    def test_denetim(self):
        h, n, t = veri()
        self.assertEqual(agent.denetle(taslak(), h, n, t), [])
        s = agent.denetle(taslak(False), h, n, t)
        for parca in ("geçersiz ['T9']", "15.10.2026 yatış-çıkış", "kesinlik ifadesi", "'K37'", "olmayan sayı: 1.5", "T1: referans dışı", "T2:", "T3:", "T7:"):
            self.assertTrue(any(parca in x for x in s), (parca, s))
        self.assertEqual(len(s), 10, s)


class AgentTesti(unittest.TestCase):
    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_dongusu(self, _):
        MESAJLAR.clear()
        with tempfile.TemporaryDirectory() as tmp:
            cikti = Path(tmp) / "e.xlsx"
            s = agent.calistir(HASTA, NOT, TET, cikti, evet=True)
            self.assertEqual((len(MESAJLAR), s["kalan"], s["tur"]), (2, [], 1))
            for gizli in ("Ayşe Deneme", "10000000146", "2026-209931", "14.06.1988"):
                self.assertNotIn(gizli, "".join(MESAJLAR))
            m = cikti.with_suffix(".txt").read_text(encoding="utf-8")
            self.assertTrue(m.startswith("EPİKRİZ (TASLAK — HEKİM ONAYI ZORUNLUDUR)"))
            self.assertIn("Yatış süresi: 3 gün", m)
            self.assertIn("  - Amoksisilin-klavulanik asit 1000 mg tablet, 2x1, 5 gün", m)
            self.assertIn("Lökosit (WBC): 16,8 10^3/µL (ref. 4-10) [Yüksek]", m)
            self.assertEqual(load_workbook(cikti)["Tetkikler"]["G2"].value, "Yüksek")

    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_kapali_ve_cli(self, _):
        with tempfile.TemporaryDirectory() as tmp:
            s = agent.calistir(HASTA, NOT, TET, Path(tmp) / "e.xlsx", tur=0, evet=True)
            self.assertEqual(len(s["kalan"]), 10)
            self.assertEqual(agent.main(["--evet", "--cikti", str(Path(tmp) / "c.xlsx")]), 0)
            bozuk = Path(tmp) / "h.txt"
            bozuk.write_text("Yatış Tarihi: 14.10.2026\nÇıkış Tarihi: 11.10.2026\n", encoding="utf-8")
            self.assertEqual(agent.main(["--evet", "--hasta", str(bozuk)]), 1)


if __name__ == "__main__":
    unittest.main()
