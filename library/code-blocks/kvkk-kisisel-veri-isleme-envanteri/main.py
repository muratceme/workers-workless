"""
KVKK Kişisel Veri İşleme Envanteri — Workers / Workless kod bloğu
Bilgi Teknolojileri › Bilgi Güvenliği Uzmanı

Departmanlardan toplanan veri işleme formlarını tek bir kişisel veri işleme envanterinde birleştirir ve eksik /
tutarsız kayıtları işaretler:
  - Hukuki sebep normalleştirilir: KVKK md. 5/1 açık rıza; md. 5/2 (a) kanunlarda açıkça öngörülme, (b) fiili
    imkânsızlık, (c) sözleşmenin kurulması / ifası, (ç) hukuki yükümlülük, (d) alenileştirme, (e) hakkın tesisi,
    kullanılması, korunması, (f) meşru menfaat.
  - Özel nitelikli veri (md. 6: sağlık, biyometrik, genetik, ceza mahkûmiyeti, din, sendika üyeliği, ırk / etnik köken,
    siyasi düşünce, felsefi inanç, kılık kıyafet, cinsel hayat) anahtar kelimeyle tespit edilir. Formdaki işaretle
    çelişki ve md. 6'da yer almayan sebebe (meşru menfaat, sözleşme) dayanma işaretlenir.
  - Yurt dışı aktarım (md. 9): ülke ve dayanak zorunlu; açık rızanın yalnız arızi aktarımda dayanak olabileceği,
    standart sözleşmenin Kurula bildirimi hatırlatılır.
  - Saklama süresi belirsiz ("süresiz", "gerektiği kadar") veya dayanaksız; teknik / idari tedbir boş; pazarlama
    sürecinde TCKN gibi ölçülülük şüphesi; açık rızanın sözleşme / yükümlülük süreçlerinde gereksiz kullanımı;
    mükerrer kayıt; formu gelmeyen departman.
Rapor: envanter, özel nitelikli veriler, yurt dışı aktarımlar, saklama ve imha tablosu, kategori özeti, departman
durumu, bulgular. İnternete bağlanmaz.

Kullanım:
    python main.py                                                 # örnek: 15 kayıt, 6 departman
    python main.py --formlar formlar.xlsx --departmanlar departmanlar.csv
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")

SUTUNLAR = {"departman": ("departman", "birim"), "surec": ("surec", "faaliyet", "is sureci"), "kisi": ("veri konusu kisi grubu", "ilgili kisi grubu", "kisi grubu"),
            "kategori": ("veri kategorisi", "kategori"), "alanlar": ("veri alanlari", "kisisel veriler", "veriler"), "ozel": ("ozel nitelikli",),
            "amac": ("isleme amaci", "amac"), "sebep": ("hukuki sebep", "isleme sarti", "hukuki dayanak"), "yontem": ("toplama yontemi",),
            "ortam": ("saklama ortami", "ortam"), "sure": ("saklama suresi",), "sure_dayanak": ("sure dayanagi", "saklama suresi dayanagi"),
            "alici": ("alici gruplari", "aktarilan alicilar", "alicilar"), "yurtdisi": ("yurt disi aktarim", "yurtdisi aktarim"),
            "ulke": ("aktarim ulkesi", "ulke"), "aktarim_dayanak": ("aktarim dayanagi",), "teknik": ("teknik tedbirler", "teknik tedbir"),
            "idari": ("idari tedbirler", "idari tedbir")}
# (anahtar kelime, normal ad, madde)
SEBEPLER = [("acik riza", "Açık rıza", "5/1"), ("kanunlarda acikca", "Kanunlarda açıkça öngörülmesi", "5/2-a"), ("kanun", "Kanunlarda açıkça öngörülmesi", "5/2-a"),
            ("fiili imkansiz", "Fiili imkânsızlık", "5/2-b"), ("sozlesme", "Sözleşmenin kurulması veya ifası", "5/2-c"),
            ("hukuki yukumluluk", "Hukuki yükümlülük", "5/2-ç"), ("alenilestir", "Alenileştirme", "5/2-d"),
            ("hakkin tesisi", "Bir hakkın tesisi, kullanılması veya korunması", "5/2-e"), ("hak tesisi", "Bir hakkın tesisi, kullanılması veya korunması", "5/2-e"),
            ("mesru menfaat", "Meşru menfaat", "5/2-f")]
OZEL_UYGUN = {"Açık rıza", "Kanunlarda açıkça öngörülmesi", "Fiili imkânsızlık", "Alenileştirme", "Bir hakkın tesisi, kullanılması veya korunması", "Hukuki yükümlülük"}
OZEL_KELIMELER = {"saglik": "Sağlık", "saglik raporu": "Sağlık", "hastalik": "Sağlık", "engelli": "Sağlık", "kan grubu": "Sağlık", "ilac": "Sağlık", "tani": "Sağlık",
                  "biyometrik": "Biyometrik", "parmak izi": "Biyometrik", "yuz tanima": "Biyometrik", "avuc ici": "Biyometrik", "retina": "Biyometrik",
                  "genetik": "Genetik", "dna": "Genetik", "adli sicil": "Ceza mahkûmiyeti", "sabika": "Ceza mahkûmiyeti", "ceza mahkumiyet": "Ceza mahkûmiyeti",
                  "din": "Din / inanç", "mezhep": "Din / inanç", "sendika": "Sendika üyeliği", "dernek uyeligi": "Dernek / vakıf üyeliği",
                  "irk": "Irk / etnik köken", "etnik": "Irk / etnik köken", "siyasi": "Siyasi düşünce", "felsefi": "Felsefi inanç", "kilik kiyafet": "Kılık kıyafet",
                  "cinsel": "Cinsel hayat"}
YURTDISI_DAYANAK = [("yeterlilik", "Yeterlilik kararı"), ("standart sozlesme", "Standart sözleşme"), ("baglayici sirket", "Bağlayıcı şirket kuralları"),
                    ("taahhutname", "Taahhütname (Kurul izni)"), ("arizi", "Arızi aktarım"), ("acik riza", "Açık rıza")]
BELIRSIZ_SURE = ("suresiz", "gerektigi kadar", "gerekli sure", "belirsiz", "surekli", "her zaman")
ZORUNLU = [("surec", "Süreç"), ("kisi", "Veri konusu kişi grubu"), ("kategori", "Veri kategorisi"), ("amac", "İşleme amacı"), ("sebep", "Hukuki sebep"),
           ("sure", "Saklama süresi")]


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(s or "").translate(_TR).lower()).split())


def metin(x) -> str:
    return str(x if x is not None else "").strip()


def bos(x) -> bool:
    return katla(x) in ("", "yok", "bos")


def tablo_oku(yol: Path) -> list[list]:
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, data_only=True, read_only=True)
        try:
            satirlar = [list(r) for r in wb.active.iter_rows(values_only=True)]
        finally:
            wb.close()
    else:
        icerik = None
        for kod in ("utf-8-sig", "cp1254"):
            try:
                icerik = yol.read_text(encoding=kod)
                break
            except UnicodeDecodeError:
                continue
        ilk = "\n".join(icerik.splitlines()[:10])
        satirlar = list(csv.reader(icerik.splitlines(), delimiter=max(";\t", key=ilk.count)))
    return [r for r in satirlar if any(c not in (None, "") for c in r)]


def kayitlar(yol: Path, sozluk: dict, zorunlu: tuple[str, ...]) -> list[dict]:
    s = tablo_oku(yol)
    for bi, r in enumerate(s[:15]):
        b = [katla(c) for c in r]
        es = {}
        for alan, adlar in sozluk.items():
            j = next((b.index(a) for a in adlar if a in b and b.index(a) not in es.values()), None)
            if j is not None:
                es[alan] = j
        if all(z in es for z in zorunlu):
            return [{a: (r[j] if j < len(r) else None) for a, j in es.items()} | {"_satir": n} for n, r in enumerate(s[bi + 1:], bi + 2)]
    raise ValueError(f"{yol.name}: başlık satırı bulunamadı; gerekli sütunlar: {', '.join(zorunlu)}")


# ----------------------------------------------------------------------------
# Kayıt
# ----------------------------------------------------------------------------

@dataclass
class Kayit:
    no: int
    v: dict
    sebepler: list = field(default_factory=list)       # [(ad, madde)]
    ozel_turler: list = field(default_factory=list)
    yurtdisi: bool = False
    dayanak: str = ""
    bulgular: list = field(default_factory=list)

    def __getitem__(self, k):
        return metin(self.v.get(k))


def sebep_coz(x: str) -> tuple[list, list]:
    """'Sözleşmenin ifası, hukuki yükümlülük' → [(ad, madde)], tanınmayan parçalar"""
    bulunan, tanınmayan = [], []
    for parca in re.split(r"[,;/+]| ve ", metin(x)):
        k = katla(parca)
        if not k:
            continue
        e = next(((ad, m) for kel, ad, m in SEBEPLER if kel in k), None)
        if e and e not in bulunan:
            bulunan.append(e)
        elif not e:
            tanınmayan.append(parca.strip())
    return bulunan, tanınmayan


def ozel_bul(*metinler: str) -> list[str]:
    k = " " + " ".join(katla(m) for m in metinler) + " "
    return sorted({tur for kel, tur in OZEL_KELIMELER.items() if f" {kel}" in k and (len(kel) > 4 or f" {kel} " in k)})


def denetle(kayitlar_: list[Kayit]) -> None:
    imza = Counter()
    for k in kayitlar_:
        imza[tuple(katla(k[a]) for a in ("departman", "surec", "kisi", "kategori", "alanlar", "amac"))] += 1
    goruldu = set()
    for k in kayitlar_:
        b = k.bulgular
        for alan, ad in ZORUNLU:
            if bos(k[alan]):
                b.append(("Yüksek" if alan in ("sebep", "amac") else "Orta", "Eksik alan", f"'{ad}' boş"))
        k.sebepler, tanınmayan = sebep_coz(k["sebep"])
        if tanınmayan:
            b.append(("Orta", "Tanınmayan hukuki sebep", f"'{', '.join(tanınmayan)}' KVKK md. 5 / 6 şartlarıyla eşleşmedi"))
        k.ozel_turler = ozel_bul(k["kategori"], k["alanlar"])
        isaret = katla(k["ozel"]).startswith(("e", "var", "1"))
        if k.ozel_turler and not isaret:
            b.append(("Orta", "Özel nitelik işareti", f"{', '.join(k.ozel_turler)} verisi tespit edildi ama 'Özel Nitelikli' Hayır işaretli"))
        if (k.ozel_turler or isaret) and k.sebepler:
            uygunsuz = [ad for ad, _ in k.sebepler if ad not in OZEL_UYGUN]
            if uygunsuz and not any(ad in OZEL_UYGUN for ad, _ in k.sebepler):
                oneri = " Çalışanların iş sağlığı ve güvenliği verilerinde md. 6/3-(e) (istihdam, İSG, sosyal güvenlik hukuki yükümlülüğü) değerlendirilebilir." \
                    if "Sağlık" in k.ozel_turler and katla(k["kisi"]).startswith("calisan") else ""
                b.append(("Yüksek", "Özel nitelikli veri sebebi", f"Özel nitelikli veri ({', '.join(k.ozel_turler) or 'işaretli'}) '{', '.join(uygunsuz)}' "
                          f"sebebine dayandırılmış; KVKK md. 6'daki şartlardan biri gerekir.{oneri}"))
            if "Biyometrik" in k.ozel_turler:
                b.append(("Orta", "Biyometrik veri", "Biyometrik verinin amaçla ölçülü olduğu ve daha az müdahaleci bir yöntemin (ör. kart) "
                          "bulunmadığı değerlendirilmelidir"))
        if ("Açık rıza", "5/1") in k.sebepler and len(k.sebepler) == 1 and not k.ozel_turler and \
                any(x in katla(k["amac"] + " " + k["surec"]) for x in ("odeme", "fatura", "bordro", "ucret", "sozlesme", "sgk", "ozluk")):
            b.append(("Bilgi", "Gereksiz açık rıza", "Süreç sözleşmenin ifası veya hukuki yükümlülüğe dayanabiliyor; diğer şart varken açık rıza istenmesi "
                      "önerilmez (geri alınabilir)"))
        if any(x in katla(k["alanlar"]) for x in ("tckn", "tc kimlik")) and any(x in katla(k["amac"] + " " + k["surec"] + " " + k["kategori"])
                                                                                for x in ("pazarlama", "kampanya", "bulten", "reklam", "memnuniyet")):
            b.append(("Orta", "Ölçülülük", "Pazarlama / iletişim amacı için TCKN toplanması amaçla sınırlı ve ölçülü olma ilkesine aykırı olabilir"))
        sure = katla(k["sure"])
        if sure and any(x in sure for x in BELIRSIZ_SURE):
            b.append(("Orta", "Belirsiz saklama süresi", f"'{k['sure']}' belirli bir süre değil; mevzuat veya şirket politikasına dayalı süre yazılmalı"))
        elif sure and bos(k["sure_dayanak"]):
            b.append(("Bilgi", "Süre dayanağı yok", "Saklama süresinin dayanağı (kanun, sözleşme, politika) yazılmalı"))
        k.yurtdisi = katla(k["yurtdisi"]).startswith(("e", "var", "1"))
        if k.yurtdisi:
            if bos(k["ulke"]):
                b.append(("Yüksek", "Yurt dışı ülke yok", "Yurt dışına aktarım var ama ülke yazılmamış"))
            d = next((ad for kel, ad in YURTDISI_DAYANAK if kel in katla(k["aktarim_dayanak"])), "")
            k.dayanak = d
            if not d:
                b.append(("Yüksek", "Yurt dışı aktarım dayanağı", f"{k['ulke'] or 'Yurt dışına'} aktarımın KVKK md. 9 dayanağı yok (yeterlilik kararı, standart "
                          "sözleşme, bağlayıcı şirket kuralları, taahhütname veya arızi aktarım)"))
            elif d == "Açık rıza":
                b.append(("Orta", "Yurt dışı aktarımda açık rıza", "Açık rıza yalnız arızi (düzenli olmayan) aktarımlarda dayanak olabilir; sürekli aktarımda "
                          "standart sözleşme gibi uygun güvence gerekir (KVKK md. 9, 2024 değişikliği)"))
            elif d == "Standart sözleşme":
                b.append(("Bilgi", "Standart sözleşme bildirimi", "Standart sözleşme imzadan itibaren 5 iş günü içinde Kurula bildirilmelidir"))
        if bos(k["teknik"]):
            b.append(("Orta", "Teknik tedbir yok", "Teknik tedbirler boş"))
        if bos(k["idari"]):
            b.append(("Orta", "İdari tedbir yok", "İdari tedbirler boş"))
        a = tuple(katla(k[x]) for x in ("departman", "surec", "kisi", "kategori", "alanlar", "amac"))
        if imza[a] > 1:
            if a in goruldu:
                b.append(("Bilgi", "Mükerrer kayıt", "Aynı süreç / kategori / amaç kaydı tekrar ediyor; envantere bir kez alındı"))
            goruldu.add(a)


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
RENK = {"Yüksek": "FDE2E1", "Orta": "FFF4CE", "Bilgi": "E8F0FE"}
KONTROL = PatternFill("solid", fgColor="FFF4CE")
UST = Alignment(vertical="top", wrap_text=True)


def _baslik(ws, basliklar, genislik=()):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "C2"


def rapor_yaz(cikti: Path, s: dict) -> None:
    kayitlar_ = [k for k in s["kayitlar"] if not any(t == "Mükerrer kayıt" for _, t, _ in k.bulgular)]
    wb = Workbook()
    en = wb.active
    en.title = "Envanter"
    _baslik(en, ["No", "Departman", "Süreç", "Veri Konusu Kişi Grubu", "Veri Kategorisi", "Veri Alanları", "Özel Nitelikli", "İşleme Amacı", "Hukuki Sebep",
                 "KVKK Maddesi", "Toplama Yöntemi", "Saklama Ortamı", "Saklama Süresi", "Süre Dayanağı", "Alıcı Grupları", "Yurt Dışı Aktarım", "Teknik Tedbirler",
                 "İdari Tedbirler", "Bulgu", "Durum"],
            (5, 15, 22, 16, 16, 28, 14, 28, 26, 10, 16, 16, 16, 20, 20, 18, 22, 22, 9, 12))
    for k in kayitlar_:
        ozel = ", ".join(k.ozel_turler) if k.ozel_turler else ("Evet" if katla(k["ozel"]).startswith("e") else "Hayır")
        yd = f"Evet · {k['ulke']} · {k.dayanak or '—'}" if k.yurtdisi else "Hayır"
        n = sum(1 for o, _, _ in k.bulgular if o != "Bilgi")
        en.append([k.no, k["departman"], k["surec"], k["kisi"], k["kategori"], k["alanlar"], ozel, k["amac"],
                   "; ".join(ad for ad, _ in k.sebepler) or k["sebep"], ", ".join(m for _, m in k.sebepler), k["yontem"], k["ortam"], k["sure"],
                   k["sure_dayanak"], k["alici"], yd, k["teknik"], k["idari"], n, ""])
        r = en.max_row
        if k.ozel_turler:
            en.cell(r, 7).fill = PatternFill("solid", fgColor="FFF4CE")
        if n:
            en.cell(r, 19).fill = PatternFill("solid", fgColor="FDE2E1" if any(o == "Yüksek" for o, _, _ in k.bulgular) else "FFF4CE")
        en.cell(r, 20).fill = KONTROL
        for c in en[r]:
            c.alignment = UST
    en.auto_filter.ref = f"A1:T{en.max_row}"

    oz = wb.create_sheet("Özel Nitelikli Veriler")
    _baslik(oz, ["No", "Departman", "Süreç", "Kişi Grubu", "Tespit Edilen Tür", "Veri Alanları", "Hukuki Sebep", "Uygun mu?"], (5, 15, 22, 16, 22, 28, 30, 30))
    for k in kayitlar_:
        if k.ozel_turler or katla(k["ozel"]).startswith("e"):
            uygun = any(ad in OZEL_UYGUN for ad, _ in k.sebepler)
            oz.append([k.no, k["departman"], k["surec"], k["kisi"], ", ".join(k.ozel_turler) or "İşaretli", k["alanlar"], "; ".join(ad for ad, _ in k.sebepler),
                       "Md. 6 şartlarından biri" if uygun else "Hayır — md. 6 şartı gerekli"])
            oz.cell(oz.max_row, 8).fill = PatternFill("solid", fgColor="E3F4E1" if uygun else "FDE2E1")

    yd = wb.create_sheet("Yurt Dışı Aktarımlar")
    _baslik(yd, ["No", "Departman", "Süreç", "Veri Kategorisi", "Alıcı", "Ülke", "Dayanak", "Not"], (5, 15, 22, 18, 24, 12, 24, 50))
    for k in kayitlar_:
        if k.yurtdisi:
            yd.append([k.no, k["departman"], k["surec"], k["kategori"], k["alici"], k["ulke"], k.dayanak or "—",
                       "; ".join(a for _, t, a in k.bulgular if t.startswith(("Yurt dışı", "Standart")))])
            yd.cell(yd.max_row, 8).alignment = UST

    si = wb.create_sheet("Saklama ve İmha")
    _baslik(si, ["Veri Kategorisi", "Süreç", "Departman", "Saklama Ortamı", "Saklama Süresi", "Dayanak", "İmha Yöntemi", "Sorumlu"], (18, 22, 15, 18, 22, 26, 18, 16))
    for k in sorted(kayitlar_, key=lambda k: (katla(k["kategori"]), k["departman"])):
        si.append([k["kategori"], k["surec"], k["departman"], k["ortam"], k["sure"], k["sure_dayanak"], "", s["sorumlu"].get(katla(k["departman"]), "")])
        si.cell(si.max_row, 7).fill = KONTROL
    si.append([])
    si.append(["Kişisel verilerin silinmesi, yok edilmesi veya anonim hâle getirilmesi Yönetmeliğine göre periyodik imha en fazla 6 ayda bir yapılır; "
               "imha yöntemleri saklama ve imha politikasında belirtilmelidir."])

    ko = wb.create_sheet("Kategori Özeti")
    _baslik(ko, ["Veri Kategorisi", "Kişi Grupları", "Amaçlar", "Hukuki Sebepler", "Departmanlar", "Kayıt"], (20, 30, 50, 40, 30, 7))
    g = defaultdict(list)
    for k in kayitlar_:
        g[k["kategori"] or "—"].append(k)
    for kat, lst in sorted(g.items()):
        ko.append([kat, ", ".join(sorted({k["kisi"] for k in lst})), "; ".join(sorted({k["amac"] for k in lst})),
                   "; ".join(sorted({ad for k in lst for ad, _ in k.sebepler})), ", ".join(sorted({k["departman"] for k in lst})), len(lst)])
        for c in ko[ko.max_row]:
            c.alignment = UST

    dp = wb.create_sheet("Departman Durumu")
    _baslik(dp, ["Departman", "Sorumlu", "Kayıt", "Özel Nitelikli", "Yurt Dışı", "Yüksek Bulgu", "Orta Bulgu", "Durum"], (18, 22, 7, 12, 10, 12, 10, 18))
    for d in s["departmanlar"]:
        lst = [k for k in kayitlar_ if katla(k["departman"]) == katla(d)]
        y = sum(1 for k in lst for o, _, _ in k.bulgular if o == "Yüksek")
        o_ = sum(1 for k in lst for o, _, _ in k.bulgular if o == "Orta")
        durum = "Form gelmedi" if not lst else "Düzeltme gerekli" if y else "Gözden geçir" if o_ else "Tamam"
        dp.append([d, s["sorumlu"].get(katla(d), ""), len(lst), sum(1 for k in lst if k.ozel_turler), sum(1 for k in lst if k.yurtdisi), y, o_, durum])
        dp.cell(dp.max_row, 8).fill = PatternFill("solid", fgColor={"Tamam": "E3F4E1", "Gözden geçir": "FFF4CE"}.get(durum, "FDE2E1"))

    bl = wb.create_sheet("Bulgular", 0)
    _baslik(bl, ["Önem", "Tür", "Kayıt No", "Departman", "Süreç", "Açıklama", "Düzeltildi"], (9, 26, 8, 15, 22, 90, 11))
    bl.freeze_panes = "A2"
    for b in s["bulgular"]:
        bl.append([b["onem"], b["tur"], b["no"], b["departman"], b["surec"], b["aciklama"], ""])
        bl.cell(bl.max_row, 1).fill = PatternFill("solid", fgColor=RENK[b["onem"]])
        bl.cell(bl.max_row, 6).alignment = UST
        bl.cell(bl.max_row, 7).fill = KONTROL
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(form_yolu: Path, cikti: Path, dep_yolu: Path | None = None) -> dict:
    ks = [Kayit(i, r) for i, r in enumerate(kayitlar(form_yolu, SUTUNLAR, ("departman", "kategori")), 1)
          if any(metin(r.get(a)) for a in ("surec", "kategori", "alanlar"))]
    if not ks:
        raise ValueError(f"{form_yolu.name}: kayıt bulunamadı")
    denetle(ks)
    sorumlu, departmanlar = {}, []
    if dep_yolu:
        for r in tablo_oku(dep_yolu)[1:]:
            if r and metin(r[0]):
                departmanlar.append(metin(r[0]))
                sorumlu[katla(r[0])] = metin(r[1]) if len(r) > 1 else ""
    for d in dict.fromkeys(k["departman"] for k in ks):
        if katla(d) not in {katla(x) for x in departmanlar}:
            departmanlar.append(d)
    bulgular = [{"onem": o, "tur": t, "no": k.no, "departman": k["departman"], "surec": k["surec"], "aciklama": a} for k in ks for o, t, a in k.bulgular]
    for d in departmanlar:
        if not any(katla(k["departman"]) == katla(d) for k in ks):
            bulgular.append({"onem": "Orta", "tur": "Form gelmedi", "no": "", "departman": d, "surec": "", "aciklama": "Departmandan veri işleme formu gelmedi"})
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    bulgular.sort(key=lambda b: (sira[b["onem"]], b["tur"], str(b["no"]).zfill(4)))
    s = {"kayitlar": ks, "bulgular": bulgular, "departmanlar": departmanlar, "sorumlu": sorumlu}
    rapor_yaz(cikti, s)
    return s


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Departman veri işleme formlarından KVKK kişisel veri işleme envanteri oluşturur, eksik ve tutarsız kayıtları işaretler.")
    p.add_argument("--formlar", type=Path, default=ORNEK / "departman_formlari.csv",
                   help="Departman, Süreç, Veri Konusu Kişi Grubu, Veri Kategorisi, Veri Alanları, Özel Nitelikli, İşleme Amacı, Hukuki Sebep, …")
    p.add_argument("--departmanlar", type=Path, help="İsteğe bağlı: Departman, Sorumlu (formu gelmeyenleri bulmak için)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "kvkk_envanter.xlsx")
    a = p.parse_args(argv)
    dep = a.departmanlar or (ORNEK / "departmanlar.csv" if a.formlar == ORNEK / "departman_formlari.csv" else None)
    for y in (a.formlar, dep):
        if y and not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    try:
        s = calistir(a.formlar, a.cikti, dep)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    ks = s["kayitlar"]
    print(f"[OK] {len(ks)} kayıt · {len(s['departmanlar'])} departman · özel nitelikli {sum(1 for k in ks if k.ozel_turler)} · "
          f"yurt dışı aktarım {sum(1 for k in ks if k.yurtdisi)}")
    for b in s["bulgular"]:
        if b["onem"] == "Yüksek":
            print(f"[!] #{b['no']} {b['departman']} · {b['surec']} · {b['tur']}: {b['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
