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

ORNEK = KLASOR / "ornek_veri"
MESAJLAR = []

# Sahte model bilerek hatalar yapar: uydurma sayı (Q-3), teslim sözü (Q-4), iletişim bilgisi (Q-5), verilmeyen kaynak (Q-7)
CEVAPLAR = {
    "Q-1": ("tam", "Merhaba, nevresim 200x220 cm ölçüsündedir, çift kişilik yataklar içindir. Teşekkürler.", ["Set İçeriği"]),
    "Q-2": ("tam", "Merhaba, ürün %100 pamuktur ve 40 °C'de yıkanabilir. Teşekkürler.", ["Malzeme", "Yıkama"]),
    "Q-3": ("kismi", "Merhaba, havlu 500 g/m² ağırlığındadır; ilk yıkamada 30 gr kadar tüy bırakabilir.", ["Malzeme"]),
    "Q-4": ("tam", "Merhaba, kupa mikrodalgaya uygundur. Bugün sipariş verirseniz yarın elinizde olur.", ["Ek Bilgi", "K1"]),
    "Q-5": ("tam", "Merhaba, bize 0532 111 22 33 numarasından WhatsApp ile ulaşabilirsiniz.", []),
    "Q-6": ("yok", "Merhaba, kontrol edip dönüş yapacağız.", []),
    "Q-7": ("tam", "Merhaba, hijyen nedeniyle ambalajı açılmış havlu iade alınamamaktadır. Teşekkürler.", ["K2", "K9"]),
}


def sahte_json_iste(sistem, kullanici, sema, max_tokens=16000):
    MESAJLAR.append(kullanici)
    idler = re.findall(r'<soru id="([^"]+)"', kullanici)
    return {"cevaplar": [{"id": i, "kapsam": CEVAPLAR[i][0], "cevap": CEVAPLAR[i][1], "dayanak": list(CEVAPLAR[i][2]),
                          "insan_gerekli": False, "gerekce": ""} for i in idler]}


class AgentTesti(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        MESAJLAR.clear()
        with mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste), tempfile.TemporaryDirectory() as tmp:
            cls.s = agent.calistir(ORNEK / "sorular.csv", ORNEK / "urunler.csv", ORNEK / "politikalar.md", Path(tmp) / "c.xlsx",
                                   evet=True)
        cls.r, cls.n = cls.s["sonuc"], cls.s["notlar"]
        cls.q = {x["id"]: x for x in cls.s["sorular"]}

    def test_temiz_cevaplar(self):
        for i in ("Q-1", "Q-2"):
            self.assertFalse(self.r[i]["insan_gerekli"], (i, self.n[i]))

    def test_uydurma_sayi(self):
        self.assertTrue(self.r["Q-3"]["insan_gerekli"])
        self.assertTrue(any("30 gr" in x for x in self.n["Q-3"]))

    def test_teslim_sozu_ve_iletisim(self):
        self.assertTrue(any("teslim" in x for x in self.n["Q-4"]))
        self.assertTrue(any("iletişim" in x for x in self.n["Q-5"]))
        self.assertIn("İletişim bilgisi istiyor", self.q["Q-5"]["etiketler"])

    def test_urun_yok_ve_dayanak_suzme(self):
        self.assertIn("Ürün kartı yok", self.q["Q-6"]["etiketler"])
        self.assertTrue(self.r["Q-6"]["insan_gerekli"])
        self.assertNotIn("K9", self.r["Q-7"]["dayanak"])
        self.assertIn("<urun_bulunamadi/>", MESAJLAR[0].split('<soru id="Q-6"')[1].split("</soru>")[0])

    def test_politika_eslesmesi(self):
        bolumler = {b.id: b for b in agent.sss.bilgi_bankasi_oku(ORNEK / "politikalar.md")}
        self.assertIn("İade", bolumler[self.q["Q-7"]["eslesme"][0][0]].baslik)


if __name__ == "__main__":
    unittest.main()
