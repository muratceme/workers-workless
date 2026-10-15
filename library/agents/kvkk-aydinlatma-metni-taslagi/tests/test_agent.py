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
ENV, SOR = V / "veri_envanteri.csv", V / "veri_sorumlusu.txt"
MESAJLAR = []


def b(baslik, metin, *d):
    return {"baslik": baslik, "metin": metin, "dayanak": list(d)}


def taslak(dogru=True) -> dict:
    t = {"giris": "Bu metin, internet sitemizin üyeleri ve müşterilerimizi kişisel verilerinin işlenmesi hakkında bilgilendirmek amacıyla hazırlanmıştır.",
         "amac_bolumleri": [
             b("Üyelik ve sipariş", "Ad, soyad, telefon numarası, e-posta ve teslimat adresiniz; üyelik ve sipariş süreçlerinin yürütülmesi, siparişin "
                                    "teslim edilmesi ve sipariş durumu hakkında bilgilendirme amaçlarıyla, KVKK md. 5/2-c uyarınca bir sözleşmenin "
                                    "kurulması veya ifasıyla doğrudan doğruya ilgili olması hukuki sebebine dayanılarak işlenir.", "E1", "E2"),
             b("Fatura ve saklama", "Sipariş, iade ve fatura bilgileriniz; fatura düzenlenmesi ve ticari defter ve belgelerin saklanması amacıyla, "
                                    "md. 5/2-ç uyarınca hukuki yükümlülüğümüzü yerine getirebilmemiz için işlenir.", "E3"),
             b("Bilgi güvenliği", "IP adresiniz, oturum ve giriş kayıtlarınız; bilgi güvenliğinin sağlanması ve hesapların yetkisiz erişime karşı "
                                  "korunması amacıyla, temel hak ve özgürlüklerinize zarar vermemek kaydıyla meşru menfaatimiz (md. 5/2-f) için işlenir.", "E4"),
             b("Pazarlama", "E-posta adresiniz ve alışveriş geçmişiniz, kişiye özel kampanya ve ticari elektronik ileti gönderilmesi amacıyla yalnız "
                            "açık rızanız bulunması hâlinde (md. 5/1) işlenir.", "E5"),
             b("Hukuki süreçler", "Şikâyet ve talep yazışmalarınız, uyuşmazlıklarda hakların korunması amacıyla md. 5/2-e uyarınca bir hakkın tesisi, "
                                  "kullanılması veya korunması için zorunlu olması hukuki sebebine dayanılarak işlenir.", "E6")],
         "aktarim": [{"metin": "Kimlik ve iletişim bilgileriniz siparişin teslim edilmesi amacıyla kargo şirketlerine aktarılır.", "dayanak": ["E1", "E2"]},
                     {"metin": "Fatura bilgileriniz vergi yükümlülükleri için mali müşavirlik hizmeti alınan firmaya ve yetkili kamu kurum ve "
                               "kuruluşlarına aktarılır.", "dayanak": ["E3"]},
                     {"metin": "İşlem güvenliği verileriniz sunucu ve yedekleme hizmeti alınan bulut barındırma hizmeti sağlayıcısına yurt dışına, "
                               "KVKK md. 9 uyarınca standart sözleşme imzalanarak aktarılır.", "dayanak": ["E4"]},
                     {"metin": "Pazarlama verileriniz e-posta gönderim hizmeti sağlayıcısına; yazışmalarınız avukatlar ile yetkili kamu kurum ve "
                               "kuruluşlarına aktarılır.", "dayanak": ["E5", "E6"]}],
         "toplama_yontemi": "Kişisel verileriniz internet sitesindeki üyelik ve sipariş formları, internet sitesindeki sipariş formu, internet sitesindeki "
                            "ileti izni kutusu, çağrı merkezi ve e-posta aracılığıyla otomatik olmayan yollarla; işlem güvenliği verileri internet "
                            "sitesi kullanımı sırasında otomatik olarak toplanır."}
    if not dogru:
        t["amac_bolumleri"][2]["metin"] = "IP adresiniz bilgi güvenliğinin sağlanması ve hesapların korunması vb. amaçlarla açık rızanıza dayanılarak işlenir."
        del t["amac_bolumleri"][4]
        t["aktarim"][2]["metin"] = "İşlem güvenliği verileriniz bulut barındırma hizmeti sağlayıcısına aktarılır."
        t["toplama_yontemi"] += " Bu metni okudum, anladım ve onaylıyorum."
    return t


def sahte_json_iste(sistem, kullanici, sema, max_tokens=16000):
    MESAJLAR.append(kullanici)
    return taslak("<duzeltme>" in kullanici)


class KodTesti(unittest.TestCase):
    def test_sebep_ve_envanter(self):
        self.assertEqual([agent.sebep_coz(x, o) for x, o in (("5/2-c", False), ("md. 5/2 (ç)", False), ("Açık rıza", True), ("6/3-f", True),
                                                              ("5/2-z", False), ("Meşru menfaat", False))], ["5/2-c", "5/2-ç", "6/3-a", "6/3-f", None, None])
        env = agent.envanter_oku(ENV)
        self.assertEqual([r["kod"] for r in env], ["5/2-c", "5/2-c", "5/2-ç", "5/2-f", "5/1", "5/2-e"])
        self.assertEqual((env[2]["alicilar"], env[3]["yd_var"], env[0]["yd_var"]),
                         (["Mali müşavirlik hizmeti alınan firma", "Yetkili kamu kurum ve kuruluşları"], True, False))
        self.assertEqual(agent.envanter_kontrol(env), [])
        with tempfile.TemporaryDirectory() as tmp:
            y = Path(tmp) / "e.csv"
            y.write_text("Veri Kategorisi;Veriler;İşleme Amacı;Hukuki Sebep;Yurt Dışı Aktarım\nSağlık;Rapor;İzin takibi;5/2-ç;\n"
                         "Kimlik;Ad;Bordro;6/3-f;\nİletişim;E-posta;Destek;Rıza gibi;Evet\n", encoding="utf-8")
            s = agent.envanter_kontrol(agent.envanter_oku(y))
            self.assertEqual(len(s), 4, s)
            self.assertTrue(any("E1" in x and "md. 6/3" in x for x in s) and any("E2" in x and "özel nitelikli olmayan" in x for x in s)
                            and any("E3" in x and "tanınmadı" in x for x in s) and any("E3" in x and "md. 9" in x for x in s))

    def test_denetim(self):
        env = agent.envanter_oku(ENV)
        self.assertEqual(agent.denetle(taslak(), env), [])
        s = agent.denetle(taslak(False), env)
        for parca in ("E4 satırının hukuki sebebi", "E6: envanter satırı", "yurt dışına aktarılıyor", "vb", "açık rıza beyanı"):
            self.assertTrue(any(parca in x for x in s), (parca, s))
        self.assertEqual(len(s), 5, s)
        t = taslak()
        t["aktarim"][0]["metin"] = "Kimlik bilgileriniz teslimat için aktarılır."
        t["amac_bolumleri"][1]["metin"] += " Ayrıca kanunlarda açıkça öngörülmesi sebebiyle de işlenebilir."
        s = agent.denetle(t, env)
        self.assertTrue(any("'Kargo şirketleri'" in x for x in s) and any("envanterde olmayan hukuki sebep" in x for x in s), s)


class AgentTesti(unittest.TestCase):
    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_dongusu(self, _):
        MESAJLAR.clear()
        with tempfile.TemporaryDirectory() as tmp:
            cikti = Path(tmp) / "a.xlsx"
            s = agent.calistir(ENV, SOR, cikti, evet=True)
            self.assertEqual((len(MESAJLAR), s["kalan"], s["tur"]), (2, [], 1))
            for gizli in ("kep.tr", "Ataşehir", "0123456789000015"):
                self.assertNotIn(gizli, "".join(MESAJLAR))
            m = cikti.with_suffix(".txt").read_text(encoding="utf-8")
            self.assertIn("veri sorumlusu Örnek E-Ticaret A.Ş. (MERSİS No: 0123456789000015)", m)
            self.assertIn("münhasıran otomatik sistemler", m)
            self.assertIn("kayıtlı elektronik posta (KEP) ile orneketicaret@hs01.kep.tr", m)
            self.assertIn("en geç otuz gün içinde ücretsiz", m)
            self.assertEqual(load_workbook(cikti)["Envanter Kontrol"]["F5"].value.split(":")[0], "5/2-f")

    @mock.patch.object(llm, "json_iste", side_effect=sahte_json_iste)
    def test_duzeltme_kapali_ve_cli(self, _):
        with tempfile.TemporaryDirectory() as tmp:
            s = agent.calistir(ENV, SOR, Path(tmp) / "a.xlsx", tur=0, evet=True)
            self.assertEqual(len(s["kalan"]), 5)
            self.assertEqual(agent.main(["--evet", "--cikti", str(Path(tmp) / "c.xlsx")]), 0)
            self.assertEqual(agent.main(["--evet", "--sorumlu", str(Path(tmp) / "yok.txt")]), 1)


if __name__ == "__main__":
    unittest.main()
