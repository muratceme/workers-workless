"""Agent testleri gerçek API çağırmaz: llm.json_iste sahte bir fonksiyonla değiştirilir."""
import copy
import json
import os
import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal as D
from pathlib import Path
from unittest import mock

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))
os.environ.setdefault("ANTHROPIC_API_KEY", "test-anahtari")  # sahte; testler API çağırmaz

import agent  # noqa: E402
import llm  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

ORNEK = KLASOR / "ornek_veri"
TEKLIF = ORNEK / "teklifler"
CEVAP = json.loads((Path(__file__).parent / "sahte_cevaplar.json").read_text(encoding="utf-8"))
BUGUN = date(2026, 10, 16)       # cuma
MESAJLAR = []


def ornek_hatali() -> dict:
    """Sık model hataları: kalem toplamını hesaplamak, adı 'düzeltmek', aynı kodu iki kaleme vermek, olmayan kod uydurmak."""
    c = copy.deepcopy(CEVAP["ornek"])
    c["kalemler"][1]["malzeme"] = "V kayışı A-42"
    c["kalemler"][3]["talep_kodu"] = "K-03"
    c["kalemler"][0]["talep_kodu"] = "K-9"
    c["nakliye_tutar"] = "0"
    c["termin_metin"] = "7 gün"
    return c


def sahte_json_iste(sistem, kullanici, sema, max_tokens=16000):
    MESAJLAR.append(kullanici)
    if "DENEME HIRDAVAT" in kullanici:
        return copy.deepcopy(CEVAP["deneme"])
    if "KURGU TEKN" in kullanici:
        return copy.deepcopy(CEVAP["kurgu"])
    return copy.deepcopy(CEVAP["ornek"]) if "<duzeltme>" in kullanici else ornek_hatali()


class KodTesti(unittest.TestCase):
    def test_ifadeler(self):
        self.assertEqual([agent.vade_gun(x) for x in ("faturadan itibaren 60 gün", "Ödeme şekli: Peşin", "30 gün vadeli çek", "2 ay vade",
                                                      "%30 peşin, kalan 60 gün", "çek", "")], [60, 0, 30, 60, 60, None, None])
        self.assertEqual(agent.termin_gun("Sipariş onayından itibaren 5 iş günü", BUGUN), (7, "5 iş günü takvim gününe çevrildi"))
        self.assertEqual(agent.termin_gun("Teslim tarihi: 24.10.2026", BUGUN)[0], 8)
        self.assertEqual([agent.termin_gun(x, BUGUN)[0] for x in ("siparişten sonra 7 gün içinde", "2-3 hafta", "stoktan", "")], [7, 21, None, None])   # aralıkta üst sınır
        self.assertEqual([agent.garanti_ay(x) for x in ("24 ay (rulman)", "2 yıl", "garantili")], [24, 24, None])
        self.assertEqual((agent.gecerlilik("Teklif 15 gün geçerlidir", date(2026, 10, 14)), agent.gecerlilik("31.10.2026", None),
                          agent.gecerlilik("15 gün", None)), (date(2026, 10, 29), date(2026, 10, 31), None))
        self.assertEqual(agent.is_gunu_ekle(BUGUN, 1), date(2026, 10, 19))

    def test_talep_oku(self):
        t = agent.talep_oku(ORNEK / "talep.csv")
        self.assertEqual([(x["kod"], x["miktar"]) for x in t], [("K-01", D(200)), ("K-02", D(100)), ("K-03", D(10)), ("K-04", D(500))])
        self.assertEqual(agent.talep_oku(None), [])
        with tempfile.TemporaryDirectory() as tmp:
            y = Path(tmp) / "t.csv"
            y.write_text("Ürün;Adet\nA;1\n", encoding="utf-8")
            with self.assertRaises(llm.LLMHatasi):
                agent.talep_oku(y)

    def test_denetim(self):
        talep = agent.talep_oku(ORNEK / "talep.csv")
        m = llm.maskele((TEKLIF / "ornek_endustriyel_eposta.txt").read_text(encoding="utf-8"))
        self.assertEqual(agent.denetle(CEVAP["ornek"], m, talep), [])
        s = agent.denetle(ornek_hatali(), m, talep)
        self.assertEqual(len(s), 4)
        for parca in ("Kalem 2: Malzeme 'V kayışı A-42' belgede geçmiyor",
                      "talep kodu 'K-9' talep listesinde yok", "talep kodu K-03 kalem 3", "Nakliye tutarı 0 belgede yazmıyor"):
            self.assertTrue(any(parca in x for x in s), parca)
        c = copy.deepcopy(CEVAP["kurgu"]) | {"teklif_verilmeyen": ["K-01", "X"], "tedarikci": ""}
        s = agent.denetle(c, llm.maskele((TEKLIF / "kurgu_teknik_teklif.txt").read_text(encoding="utf-8")), talep)
        self.assertEqual(len(s), 3)
        self.assertEqual(agent.denetle({"teklif_var": False}, m, talep), [])

    def test_degerlendirme(self):
        talep = agent.talep_oku(ORNEK / "talep.csv")
        d = agent.degerlendir(CEVAP["ornek"], "", talep, BUGUN)
        self.assertEqual((d["vade"], d["termin"], d["garanti"], d["gecerlilik"], d["nakliye"], d["durum"], d["bulgular"]),
                         (60, 7, 12, date(2026, 10, 31), D(0), "Uygun", []))
        d = agent.degerlendir(CEVAP["kurgu"], "", talep, BUGUN)
        self.assertEqual([o for o, _ in d["bulgular"]], ["Bilgi", "Dikkat", "Bilgi", "Hata", "Bilgi"])
        self.assertEqual((d["vade"], d["nakliye"], d["durum"]), (0, D("60.00"), "Hatalı"))
        d = agent.degerlendir(CEVAP["deneme"], "", talep, BUGUN)
        self.assertEqual([o for o, _ in d["bulgular"]], ["Bilgi", "Dikkat", "Bilgi", "Hata", "Dikkat", "Bilgi"])
        d = agent.degerlendir(CEVAP["deneme"], "", talep, date(2026, 10, 9))
        self.assertIn(("Dikkat", "Teklif 10.10.2026 tarihinde geçerliliğini yitiriyor"), d["bulgular"])
        c = copy.deepcopy(CEVAP["deneme"]) | {"odeme_metin": "", "termin_metin": "", "kdv": "Belirtilmemiş"}
        c["kalemler"][0].update(talep_kodu="", miktar="")
        b = [m for _, m in agent.degerlendir(c, "", talep, BUGUN)["bulgular"]]
        self.assertTrue(any("boş vadeyi 0 gün" in x for x in b) and any("boş termini 0 gün" in x for x in b) and any("KDV dahil mi" in x for x in b)
                        and any("hiçbir malzemeyle eşleşmedi" in x for x in b))
        b = agent.degerlendir(CEVAP["deneme"] | {"termin_metin": "01.10.2026"}, "", talep, BUGUN)["bulgular"]
        self.assertIn(("Hata", "Teslim tarihi geçmiş (01.10.2026)"), b)


class AgentTesti(unittest.TestCase):
    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_ornek_duzeltme_ve_cikti(self, _):
        MESAJLAR.clear()
        with tempfile.TemporaryDirectory() as tmp:
            cikti = Path(tmp) / "t.xlsx"
            self.assertEqual(agent.main(["--evet", "--cikti", str(cikti)]), 0)
            self.assertEqual(len(MESAJLAR), 4)                       # 3 teklif + Örnek Endüstriyel için 1 düzeltme
            self.assertIn("K-01 | Rulman 6205-2RS | 200 adet", MESAJLAR[0])
            self.assertIn("<bugun>16.10.2026</bugun>", MESAJLAR[0])
            for m in MESAJLAR:
                self.assertNotRegex(m, r"0216 444|@ornekendustriyel")
            wb = load_workbook(cikti)
            self.assertEqual(wb.sheetnames, ["Teklif Listesi", "Teklifler", "Kontrol"])
            ws = wb["Teklif Listesi"]
            self.assertEqual([c.value for c in ws[1]], agent.KOD_BLOGU_SUTUNLARI)
            satir = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
            self.assertEqual(len(satir), 10)
            self.assertEqual(satir[0][:11], ["DENEME HIRDAVAT VE SANAYİ MALZEMELERİ", "K-01", "Rulman 6205-2RS", 150, "adet", 88.5, "TL", None, 8, 30, None])
            self.assertEqual(satir[3][:11], ["KURGU TEKNİK MAKİNA SAN. A.Ş.", "K-01", "Rulman 6205-2RS", 200, "adet", 2.1, "USD", 60, 7, 0, 24])
            self.assertEqual(satir[6][:11], ["Örnek Endüstriyel Ltd. Şti.", "K-01", "Rulman 6205-2RS", 200, "adet", 95, "TL", 0, 7, 60, 12])
            self.assertIn("Teklifteki ad: SKF muadili 6205 2RS rulman", satir[6][13])
            self.assertEqual([r[17] for r in wb["Teklifler"].iter_rows(min_row=2, values_only=True)], ["Hatalı", "Hatalı", "Uygun"])

    @mock.patch.object(llm, "json_iste", side_effect=lambda *a, **k: {"teklif_var": False, "notlar": ["Bu bir sipariş formu."], **{
        x: "" for x in ("tedarikci", "teklif_no", "teklif_tarihi", *agent.IFADE, "nakliye_tutar")}, "para_birimi": "TL", "kdv": "Belirtilmemiş",
        "kalemler": [], "teklif_verilmeyen": [], "sartlar": []})
    def test_teklif_degil_ve_hatalar(self, _):
        with tempfile.TemporaryDirectory() as tmp:
            y = Path(tmp) / "siparis.txt"
            y.write_text("SİPARİŞ FORMU\nRulman 6205 200 adet lütfen gönderiniz.\n", encoding="utf-8")
            s = agent.calistir(y, Path(tmp) / "t.xlsx", None, BUGUN, evet=True)
            self.assertEqual((s["teklifler"][0]["teklif"]["teklif_var"], s["teklifler"][0]["kalan"]), (False, []))
            self.assertIn("sipariş formu", [r for r in load_workbook(Path(tmp) / "t.xlsx")["Teklifler"].iter_rows(min_row=2, values_only=True)][0][18])
            self.assertEqual(agent.main(["--girdi", str(y), "--tarih", "yarın", "--evet"]), 2)
            self.assertEqual(agent.main(["--girdi", str(Path(tmp) / "yok"), "--evet"]), 1)
            self.assertEqual(agent.main(["--girdi", str(y), "--talep", str(Path(tmp) / "yok.csv"), "--evet"]), 1)


if __name__ == "__main__":
    unittest.main()
