import json
import sys
import tempfile
import unittest
from pathlib import Path

KLASOR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KLASOR))

import main  # noqa: E402

CV = KLASOR / "ornek_veri" / "cvler"
KRITER = KLASOR / "kriterler_ornek.json"
TAM_CV = """Can Örnek
Muhasebe Uzmanı
DENEYİM
Örnek A.Ş. 2016 - 2023
Deneme Ltd. 2023 - Günümüz
EĞİTİM
Ankara Üniversitesi, İktisat (Lisans) 2011 - 2015
BECERİLER: Excel, Bordro, Logo
YABANCI DİL: İngilizce (B2)
"""


def kriter_yaz(t: Path, k: dict) -> Path:
    yol = t / "k.json"
    yol.write_text(json.dumps(k, ensure_ascii=False), encoding="utf-8")
    return yol


class OrnekTesti(unittest.TestCase):
    def test_kararlar(self):
        with tempfile.TemporaryDirectory() as tmp:
            s = main.calistir(CV, KRITER, Path(tmp) / "c.xlsx")
        k = {x["aday"].dosya: x for x in s["adaylar"]}
        self.assertEqual(k["ayse_yilmaz.txt"]["karar"], "GEÇTİ")
        self.assertEqual(k["bozuk_dosya.pdf"]["karar"], "MANUEL KONTROL")       # okunamayan CV elenmez
        self.assertEqual(k["zeynep_demir.pdf"]["karar"], "ELENDİ")
        self.assertIn(("Deneyim", "✗", "4 yıl (en az 5)"), k["zeynep_demir.pdf"]["sonuc"])


class KuralTesti(unittest.TestCase):
    def test_belirsiz_deneyim_manuel(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "cv.txt").write_text(TAM_CV.replace("2016 - 2023", "").replace("2023 - Günümüz", ""), encoding="utf-8")
            s = main.calistir(t / "cv.txt", kriter_yaz(t, {"en_az_deneyim_yil": 3, "zorunlu_beceriler": ["Excel"]}), t / "c.xlsx")
        self.assertEqual(s["adaylar"][0]["karar"], "MANUEL KONTROL")

    def test_tam_eslesme_ve_grup(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            (t / "cv.txt").write_text(TAM_CV, encoding="utf-8")
            k = {"en_az_deneyim_yil": 5, "en_az_egitim": "Lisans", "zorunlu_diller": ["İngilizce"], "zorunlu_beceriler": ["Excel"],
                 "beceri_gruplari": [{"liste": ["SAP", "Logo"], "en_az": 1}]}
            s = main.calistir(t / "cv.txt", kriter_yaz(t, k), t / "c.xlsx")
        self.assertEqual(s["adaylar"][0]["karar"], "GEÇTİ")

    def test_ayrimci_kriter_reddedilir(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            for k in ({"zorunlu_ifadeler": ["bayan"]}, {"zorunlu_ifadeler": ["25-35 yaş"]}, {"zorunlu_ifadeler": ["B sınıfı ehliyet"]}):
                with self.assertRaises(SystemExit):
                    main.calistir(CV, kriter_yaz(t, k), t / "c.xlsx")
            # "dikkat" kriteri işe dayalı gerekçeyle kabul edilir
            k = {"zorunlu_ifadeler": [{"ifade": "B sınıfı ehliyet", "gerekce": "Sahada şirket aracıyla çalışılacak"}]}
            s = main.calistir(CV, kriter_yaz(t, k), t / "c.xlsx")
        self.assertEqual(len(s["adaylar"]), 4)

    def test_gecersiz_egitim(self):
        with tempfile.TemporaryDirectory() as tmp:
            t = Path(tmp)
            with self.assertRaises(SystemExit):
                main.calistir(CV, kriter_yaz(t, {"en_az_egitim": "Üniversite"}), t / "c.xlsx")


if __name__ == "__main__":
    unittest.main()
