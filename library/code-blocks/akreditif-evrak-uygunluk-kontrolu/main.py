"""
Akreditif Evrak Uygunluk Kontrolü — Workers / Workless kod bloğu
Dış Ticaret › Dış Ticaret Uzmanı

Akreditif şartlarını ibraz edilecek evrakların bilgileriyle karşılaştırır ve olası rezervleri (uyumsuzlukları) ibrazdan
önce listeler. Kontroller UCP 600 (ICC Akreditiflere İlişkin Yeknesak Usuller ve Uygulamalar) maddelerine dayanır:
  - md. 6, 14(c): ibraz, yükleme tarihinden sonraki 21 takvim günü içinde (akreditif başka süre verebilir) ve her
    hâlde akreditifin vadesinden önce yapılır. md. 29(a): vade bankanın kapalı olduğu güne gelirse izleyen ilk iş gününe
    uzar (ibraz için son gün de); son yükleme tarihi uzamaz (md. 29(c)). Kod yalnız Cumartesi / Pazar'ı bilir.
  - md. 14(d): evraklardaki bilgiler birebir aynı olmak zorunda değildir, ancak birbiriyle ve akreditifle çelişemez.
  - md. 14(i): evrak tarihi ibraz tarihinden sonra olamaz.
  - md. 18: fatura lehtarca, amir adına, akreditifin para biriminde düzenlenir; mal tanımı akreditiftekiyle uyuşmalıdır.
  - md. 20(a): konşimento yükleme (on board) kaydı, akreditifteki limanlar, tam takım asıl. md. 27: temiz taşıma belgesi.
  - md. 28: sigorta belgesi akreditifin para biriminde, en az CIF / CIP değerinin %110'u; tarihi yüklemeden sonra olamaz.
  - md. 30: "about / approximately / yaklaşık" → tutar, miktar ve birim fiyatta ±%10; koli veya adet olarak
    belirtilmeyen miktarda ±%5 (çekiş toplamı akreditif tutarını aşmamak kaydıyla).
  - md. 31: kısmi yükleme yasaklanmadıkça serbesttir.
Kod yalnız girilen alanları karşılaştırır; evrakın kendisini okumaz. Nihai inceleme bankanındır. İnternete bağlanmaz.

Kullanım:
    python main.py                                                 # örnek: 6 rezerv riski olan bir ibraz seti
    python main.py --akreditif akreditif.csv --istenen istenen_evraklar.csv --evraklar evraklar.csv --ibraz 02.11.2026
"""
from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
BASLIK_DOLGU, BASLIK_YAZI = PatternFill("solid", fgColor="1D1D1F"), Font(bold=True, color="FFFFFF")
RENK = {"Yüksek": "FDE2E1", "Orta": "FFF4CE", "Bilgi": "E8F0FE"}
KIRMIZI, SARI, MAVI, YESIL = "FDE2E1", "FFF4CE", "E8F0FE", "E3F4E1"
UST = Alignment(vertical="top", wrap_text=True)
PF, TF = "#,##0.00", "DD.MM.YYYY"
KURUS = Decimal("0.01")


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(s or "").translate(_TR).lower()).split())


def tarih(x) -> date | None:
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    s = str(x or "").strip()[:10]
    for f in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def para(x) -> Decimal | None:
    """1.250,50 · 1250.50 · 750.000 · %12,5 yazımlarını Decimal'e çevirir."""
    if x in (None, ""):
        return None
    if isinstance(x, (int, float, Decimal)):
        return Decimal(str(x))
    s = re.sub(r"[^\d,.\-]", "", str(x))
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def yuvarla(x: Decimal, hane: str = "0.01") -> Decimal:
    return x.quantize(Decimal(hane), rounding=ROUND_HALF_UP)


def metin(x) -> str:
    return str(x if x is not None else "").strip()


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
    """Başlık satırını bulur (ilk 15 satırda); sütun adlarını Türkçe karakter ve büyük-küçük harf ayırmadan eşleştirir."""
    s = tablo_oku(yol)
    for bi, r in enumerate(s[:15]):
        b = [katla(c) for c in r]
        es = {}
        for alan, adlar in sozluk.items():
            j = next((b.index(a) for a in adlar if a in b and b.index(a) not in es.values()), None)
            if j is not None:
                es[alan] = j
        if all(z in es for z in zorunlu):
            return [{a: None for a in sozluk} | {a: (r[j] if j < len(r) else None) for a, j in es.items()} | {"_satir": n}
                    for n, r in enumerate(s[bi + 1:], bi + 2)]
    raise ValueError(f"{yol.name}: başlık satırı bulunamadı; gerekli sütunlar: {', '.join(sozluk[z][0] for z in zorunlu)}")


def uyari(onem: str, tur: str, kim, aciklama: str) -> dict:
    return {"onem": onem, "tur": tur, "kim": metin(kim), "aciklama": aciklama}


def sirala(uyarilar: list[dict]) -> list[dict]:
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    return sorted(uyarilar, key=lambda u: (sira[u["onem"]], u["tur"], u["kim"]))


def _baslik(ws, basliklar, genislik=(), dondur="A2"):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = dondur


def _satir(ws, degerler, para_s=(), tarih_s=(), dolgu=None, sar=()):
    """Satır ekler; para_s / tarih_s sütunlarına biçim, dolgu = {sütun: renk}, sar = metni kaydırılacak sütunlar."""
    ws.append([float(d) if isinstance(d, Decimal) else d for d in degerler])
    n = ws.max_row
    for j in para_s:
        ws.cell(n, j).number_format = PF
    for j in tarih_s:
        ws.cell(n, j).number_format = TF
    for j, renk in (dolgu or {}).items():
        if renk:
            ws.cell(n, j).fill = PatternFill("solid", fgColor=renk)
    for j in sar:
        ws.cell(n, j).alignment = UST
    return n


def uyari_sayfasi(wb, uyarilar: list[dict]) -> None:
    ws = wb.create_sheet("Uyarılar")
    _baslik(ws, ["Önem", "Tür", "Kim / Ne", "Açıklama"], (9, 32, 28, 100))
    for u in uyarilar:
        _satir(ws, [u["onem"], u["tur"], u["kim"], u["aciklama"]], dolgu={1: RENK[u["onem"]]}, sar=(4,))


def kaydet(wb, cikti: Path) -> None:
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


# ----------------------------------------------------------------------------
# Okuma
# ----------------------------------------------------------------------------

EVRAK_TURLERI = (("sigorta", ("sigorta", "insurance")), ("fatura", ("fatura", "invoice")), ("konsimento", ("konsimento", "bill of lading", "b l", "tasima senedi", "cmr", "awb")),
                 ("ceki", ("ceki", "packing", "koli listesi")), ("mense", ("mense", "origin")), ("police", ("police", "draft", "kambiyo")))
EVRAK_ADI = {"fatura": "Ticari fatura", "konsimento": "Konşimento / taşıma belgesi", "sigorta": "Sigorta belgesi", "ceki": "Çeki listesi",
             "mense": "Menşe şahadetnamesi", "police": "Poliçe (draft)"}
SAYIM_BIRIMLERI = {"adet", "ad", "koli", "paket", "set", "takim", "cift", "pcs", "pieces", "cartons", "ctn", "units", "unit", "pc"}
SIGORTA_TESLIM = {"CIF", "CIP"}


def evrak_turu(ad) -> str:
    k = katla(ad)
    for tur, anahtarlar in EVRAK_TURLERI:
        if any(a in k for a in anahtarlar):
            return tur
    return k


def alan_deger(yol: Path) -> dict:
    """Alan;Değer biçimindeki dosyayı {katlanmış alan: değer} sözlüğüne çevirir."""
    return {katla(r[0]): r[1] for r in tablo_oku(yol) if len(r) > 1 and metin(r[0]) and katla(r[0]) != "alan" and r[1] not in (None, "")}


def al(d: dict, *adlar):
    return next((d[a] for a in adlar if a in d), None)


def akreditif_oku(yol: Path) -> dict:
    d = alan_deger(yol)
    a = {
        "no": metin(al(d, "akreditif no", "akreditif numarasi", "lc no")), "pb": metin(al(d, "para birimi", "doviz")).upper(),
        "tutar": para(al(d, "tutar", "akreditif tutari")), "tutar_tol": al(d, "tutar toleransi", "tutar toleransi yuzde"),
        "miktar": para(al(d, "miktar")), "birim": metin(al(d, "miktar birimi", "birim")), "miktar_tol": al(d, "miktar toleransi", "miktar toleransi yuzde"),
        "son_yukleme": tarih(al(d, "son yukleme tarihi", "son yukleme")), "vade": tarih(al(d, "vade", "vade tarihi", "gecerlilik tarihi", "son ibraz tarihi")),
        "ibraz_gun": para(al(d, "ibraz suresi", "ibraz suresi gun")), "kismi": katla(al(d, "kismi yukleme", "kismi sevkiyat")),
        "lehtar": metin(al(d, "lehtar", "lehdar")), "amir": metin(al(d, "amir", "amir alici")), "mal": metin(al(d, "mal tanimi", "malin tanimi")),
        "teslim": metin(al(d, "teslim sekli")).upper().split(" ")[0] if al(d, "teslim sekli") else "",
        "yukleme_limani": metin(al(d, "yukleme limani", "yukleme yeri")), "bosaltma_limani": metin(al(d, "bosaltma limani", "varis limani", "varis yeri")),
        "sigorta_oran": para(al(d, "sigorta orani", "sigorta orani yuzde")), "mense": metin(al(d, "mense", "mense ulke")),
    }
    eksik = [ad for ad, k in (("Para Birimi", "pb"), ("Tutar", "tutar"), ("Vade", "vade"), ("Lehtar", "lehtar"), ("Amir", "amir")) if not a[k]]
    if eksik:
        raise ValueError(f"{yol.name}: zorunlu alan eksik veya okunamadı: {', '.join(eksik)}")
    return a


def tolerans(deger, yaklasik_metni: str = "") -> Decimal | None:
    """'10', '%5', '+/-10' → oran; 'about / yaklaşık' → %10 (md. 30(a))."""
    k = katla(deger)
    if any(x in k for x in ("about", "approx", "yaklasik", "circa")):
        return Decimal("10")
    return para(deger)


def istenen_oku(yol: Path) -> list[dict]:
    sut = {"evrak": ("evrak", "belge", "evrak adi"), "asil": ("asil", "asil adet"), "kopya": ("kopya", "kopya adet"), "sart": ("sart", "ozel sart", "aciklama")}
    return [{"ad": metin(r["evrak"]), "tur": evrak_turu(r["evrak"]), "asil": int(para(r["asil"]) or 0), "kopya": int(para(r["kopya"]) or 0), "sart": metin(r["sart"])}
            for r in kayitlar(yol, sut, ("evrak",)) if metin(r["evrak"])]


def evraklar_oku(yol: Path) -> dict:
    sut = {"evrak": ("evrak", "belge"), "alan": ("alan", "bilgi"), "deger": ("deger", "icerik")}
    e: dict = {}
    for r in kayitlar(yol, sut, ("evrak", "alan", "deger")):
        if metin(r["evrak"]) and metin(r["alan"]):
            e.setdefault(evrak_turu(r["evrak"]), {"_ad": metin(r["evrak"])})[katla(r["alan"])] = r["deger"]
    return e


# ----------------------------------------------------------------------------
# Kontroller
# ----------------------------------------------------------------------------

def ayni(a, b) -> bool:
    return katla(a) == katla(b)


def iceriyor(uzun, kisa) -> bool:
    """Akreditifteki ad evrakta geçiyor mu (adres ekleri sorun sayılmaz)."""
    u, k = katla(uzun), katla(kisa)
    return bool(k) and (k in u or u in k)


def is_gunu(d: date) -> date:
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def bulgu(onem, madde, evrak, konu, akreditifte, evrakta, oneri) -> dict:
    return {"onem": onem, "madde": madde, "evrak": evrak, "konu": konu, "akreditifte": akreditifte, "evrakta": evrakta, "oneri": oneri}


def g(d) -> str:
    return d.strftime("%d.%m.%Y") if d else "—"


def s(x: Decimal | None) -> str:
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if x is not None else "—"


def kontrol(a: dict, istenen: list[dict], e: dict, ibraz: date) -> dict:
    b: list[dict] = []
    fat, kon, sig = e.get("fatura", {}), e.get("konsimento", {}), e.get("sigorta", {})
    yukleme = tarih(al(kon, "yukleme tarihi", "on board tarihi", "tarih"))
    ibraz_gun = int(a["ibraz_gun"] or 21)
    vade_fiili = is_gunu(a["vade"])
    sure_sonu = is_gunu(yukleme + timedelta(days=ibraz_gun)) if yukleme else None      # md. 29(a): ibraz için son gün de uzar
    son_ibraz = min(vade_fiili, sure_sonu) if yukleme else vade_fiili

    # --- süreler (md. 6, 14(c), 29)
    if ibraz > vade_fiili:
        b.append(bulgu("Yüksek", "md. 6(d), 29(a)", "Tüm set", "Akreditifin vadesi geçmiş", g(a["vade"]), f"İbraz {g(ibraz)}",
                       "Vade geçtikten sonra ibraz reddedilir. Amirden vade uzatımı (değişiklik) isteyin."))
    elif vade_fiili != a["vade"] and ibraz > a["vade"]:
        b.append(bulgu("Bilgi", "md. 29(a)", "Tüm set", "Vade hafta sonuna geliyor; ilk iş gününe uzar", g(a["vade"]), f"İbraz {g(ibraz)}",
                       "İbraz mektubunda md. 29(a) uyarınca süre içinde ibraz edildiği belirtilmelidir (md. 29(b)). Resmî tatilleri ayrıca kontrol edin."))
    if yukleme:
        if a["son_yukleme"] and yukleme > a["son_yukleme"]:
            b.append(bulgu("Yüksek", "md. 29(c)", EVRAK_ADI["konsimento"], "Geç yükleme", f"Son yükleme {g(a['son_yukleme'])}", f"Yükleme {g(yukleme)}",
                           "Son yükleme tarihi tatil nedeniyle uzamaz. Akreditif değişikliği gerekir."))
        if sure_sonu < ibraz <= vade_fiili:
            b.append(bulgu("Yüksek", "md. 14(c)", "Tüm set", "Geç ibraz", f"Yüklemeden sonra {ibraz_gun} gün: {g(sure_sonu)}", f"İbraz {g(ibraz)}",
                           "İbraz süresi aşılmış. Amirin rezervi kabul etmesi gerekir."))
    else:
        b.append(bulgu("Orta", "md. 14(c)", EVRAK_ADI["konsimento"], "Yükleme tarihi girilmemiş", "", "", "İbraz süresi ve son yükleme kontrolü yapılamadı."))

    # --- istenen evraklar, asıl / kopya sayısı (md. 14, 17)
    liste = []
    for i in istenen:
        ev = e.get(i["tur"])
        asil, kopya = (int(para(al(ev, "asil adet", "asil")) or 0), int(para(al(ev, "kopya adet", "kopya")) or 0)) if ev else (0, 0)
        durum = "Tamam"
        if not ev:
            durum = "Eksik"
            b.append(bulgu("Yüksek", "md. 14(a), 15", i["ad"], "İstenen evrak ibraz setinde yok", f"{i['asil']} asıl, {i['kopya']} kopya", "Yok", "Evrakı temin edin."))
        elif asil < i["asil"] or kopya < i["kopya"]:
            durum = "Adet eksik"
            b.append(bulgu("Yüksek", "md. 17", i["ad"], "Asıl / kopya sayısı eksik", f"{i['asil']} asıl, {i['kopya']} kopya", f"{asil} asıl, {kopya} kopya",
                           "Eksik nüshaları tamamlayın. Her evraktan en az bir asıl ibraz edilir (md. 17(a))."))
        liste.append({**i, "ibraz_asil": asil, "ibraz_kopya": kopya, "durum": durum})
    fazla = [v["_ad"] for t, v in e.items() if t not in {i["tur"] for i in istenen}]
    if fazla and istenen:
        b.append(bulgu("Bilgi", "md. 14(g)", "; ".join(fazla), "Akreditifte istenmeyen evrak", "İstenmiyor", "Sette var", "Banka istenmeyen evrakı dikkate almaz ve iade edebilir."))

    # --- evrak tarihleri (md. 14(i))
    for t, ev in e.items():
        d = tarih(al(ev, "tarih", "duzenleme tarihi"))
        if d and d > ibraz:
            b.append(bulgu("Yüksek", "md. 14(i)", ev["_ad"], "Evrak tarihi ibraz tarihinden sonra", f"İbraz {g(ibraz)}", g(d), "Evrak tarihi ibraz tarihini geçemez."))

    # --- fatura (md. 18, 30)
    tutar = para(al(fat, "tutar", "toplam tutar", "fatura tutari"))
    tol = tolerans(a["tutar_tol"]) or Decimal(0)
    ust = yuvarla(a["tutar"] * (1 + tol / 100))
    miktar = para(al(fat, "miktar"))
    mtol = tolerans(a["miktar_tol"])
    if fat:
        if al(fat, "para birimi") and metin(al(fat, "para birimi")).upper() != a["pb"]:
            b.append(bulgu("Yüksek", "md. 18(a)(iii)", EVRAK_ADI["fatura"], "Para birimi farklı", a["pb"], metin(al(fat, "para birimi")).upper(), "Fatura akreditifin para biriminde düzenlenir."))
        if tutar is not None and tutar > ust:
            b.append(bulgu("Yüksek", "md. 18(b), 30", EVRAK_ADI["fatura"], "Fatura tutarı akreditif tutarını aşıyor",
                           f"{s(a['tutar'])} {a['pb']}" + (f" (+%{tol} = {s(ust)})" if tol else ""), f"{s(tutar)} {a['pb']}",
                           "Aşan kısım için ödeme yapılmaz; banka faturayı reddedebilir."))
        if al(fat, "duzenleyen", "satici") and not iceriyor(al(fat, "duzenleyen", "satici"), a["lehtar"]):
            b.append(bulgu("Yüksek", "md. 18(a)(i)", EVRAK_ADI["fatura"], "Fatura lehtar tarafından düzenlenmemiş görünüyor", a["lehtar"], metin(al(fat, "duzenleyen", "satici")),
                           "Faturayı düzenleyen, akreditifteki lehtar olmalıdır."))
        if al(fat, "alici") and not iceriyor(al(fat, "alici"), a["amir"]):
            b.append(bulgu("Yüksek", "md. 18(a)(ii)", EVRAK_ADI["fatura"], "Fatura amir adına düzenlenmemiş", a["amir"], metin(al(fat, "alici")), "Fatura amir adına düzenlenir."))
        if a["mal"] and al(fat, "mal tanimi") and not ayni(al(fat, "mal tanimi"), a["mal"]):
            b.append(bulgu("Yüksek", "md. 18(c)", EVRAK_ADI["fatura"], "Mal tanımı akreditiftekiyle uyuşmuyor", a["mal"], metin(al(fat, "mal tanimi")),
                           "Faturadaki mal tanımı akreditiftekiyle uyuşmalıdır; akreditifteki ifadeyi aynen kullanın."))
        if a["teslim"] and al(fat, "teslim sekli") and metin(al(fat, "teslim sekli")).upper().split(" ")[0] != a["teslim"]:
            b.append(bulgu("Yüksek", "md. 14(d)", EVRAK_ADI["fatura"], "Teslim şekli farklı", a["teslim"], metin(al(fat, "teslim sekli")), "Teslim şekli akreditifle çelişemez."))
        if a["miktar"] and miktar is not None:
            if mtol is None:
                mtol = Decimal(0) if katla(a["birim"]) in SAYIM_BIRIMLERI else Decimal(5)
            alt_m, ust_m = a["miktar"] * (1 - mtol / 100), a["miktar"] * (1 + mtol / 100)
            if miktar > ust_m:
                b.append(bulgu("Yüksek", "md. 30", EVRAK_ADI["fatura"], "Miktar akreditiftekinden fazla", f"{a['miktar']:f} {a['birim']} (±%{mtol})", f"{miktar:f}",
                               "Fazla yükleme rezerv nedenidir."))
            elif miktar < alt_m:
                if "yasak" in a["kismi"] or "not allowed" in a["kismi"] or "prohibited" in a["kismi"] or a["kismi"] in ("hayir", "izinsiz"):
                    b.append(bulgu("Yüksek", "md. 31", EVRAK_ADI["fatura"], "Kısmi yükleme yasak, miktar eksik", f"{a['miktar']:f} {a['birim']} (±%{mtol})", f"{miktar:f}",
                                   "Akreditif kısmi yüklemeye izin vermiyor."))
                else:
                    b.append(bulgu("Bilgi", "md. 31(a)", EVRAK_ADI["fatura"], "Kısmi yükleme", f"{a['miktar']:f} {a['birim']}", f"{miktar:f}",
                                   "Kısmi yükleme serbest. Kalan miktar için son yükleme tarihini izleyin."))
    # --- konşimento (md. 20, 27)
    if kon:
        if al(kon, "serh", "kayit", "not") and not any(x in katla(al(kon, "serh", "kayit", "not")) for x in ("yok", "temiz", "clean")):
            b.append(bulgu("Yüksek", "md. 27", EVRAK_ADI["konsimento"], "Taşıma belgesi temiz değil (malın / ambalajın kusuruna ilişkin şerh var)", "Temiz", metin(al(kon, "serh", "kayit", "not")),
                           "Banka yalnız temiz taşıma belgesini kabul eder."))
        yk = katla(al(kon, "yukleme kaydi", "on board"))
        if "yukleme kaydi" in kon or "on board" in kon:
            if yk in ("yok", "hayir", "no") or "received" in yk or "teslim alindi" in yk:
                b.append(bulgu("Yüksek", "md. 20(a)(ii)", EVRAK_ADI["konsimento"], "Yükleme (on board) kaydı yok", "Shipped on board", metin(al(kon, "yukleme kaydi", "on board")),
                               "Konşimento malın adı geçen gemiye yüklendiğini göstermelidir."))
        duz, ibr = para(al(kon, "duzenlenen asil", "duzenlenen asil adet")), para(al(kon, "asil adet", "asil"))
        if duz and ibr is not None and ibr < duz:
            b.append(bulgu("Yüksek", "md. 20(a)(iv)", EVRAK_ADI["konsimento"], "Tam takım asıl konşimento ibraz edilmiyor", f"{duz:f} asıl düzenlenmiş", f"{ibr:f} asıl",
                           "Birden fazla asıl düzenlendiyse tam takım ibraz edilir (akreditif aksini söylemedikçe)."))
        for alan, ad in (("yukleme_limani", "Yükleme limanı"), ("bosaltma_limani", "Boşaltma limanı")):
            v = al(kon, katla(ad), katla(ad).replace("limani", "yeri"))
            if a[alan] and v and not iceriyor(v, a[alan]):
                b.append(bulgu("Yüksek", "md. 20(a)(iii)", EVRAK_ADI["konsimento"], f"{ad} akreditiften farklı", a[alan], metin(v), "Limanlar akreditiftekiyle aynı olmalıdır."))
    # --- sigorta (md. 28)
    if sig:
        oran = a["sigorta_oran"] or Decimal(110)
        st, sd = para(al(sig, "sigorta tutari", "tutar", "sigorta bedeli")), tarih(al(sig, "tarih", "duzenleme tarihi"))
        if al(sig, "para birimi") and metin(al(sig, "para birimi")).upper() != a["pb"]:
            b.append(bulgu("Yüksek", "md. 28(f)(i)", EVRAK_ADI["sigorta"], "Sigorta para birimi akreditiften farklı", a["pb"], metin(al(sig, "para birimi")).upper(),
                           "Sigorta teminatı akreditifin para biriminde olmalıdır."))
        if st is not None and tutar is not None and st < yuvarla(tutar * oran / 100):
            b.append(bulgu("Yüksek", "md. 28(f)(ii)", EVRAK_ADI["sigorta"], "Sigorta tutarı yetersiz", f"Fatura {s(tutar)} × %{oran:f} = {s(yuvarla(tutar * oran / 100))}", s(st),
                           "Teminat en az CIF / CIP değerinin %110'u olmalıdır (akreditif başka oran vermedikçe)."))
        bas = tarih(al(sig, "teminat baslangici", "teminat baslangic tarihi"))
        if yukleme and sd and sd > yukleme and not (bas and bas <= yukleme):
            b.append(bulgu("Yüksek", "md. 28(e)", EVRAK_ADI["sigorta"], "Sigorta tarihi yüklemeden sonra", f"Yükleme {g(yukleme)}", g(sd),
                           "Belgede teminatın en geç yükleme tarihinde başladığı yazmıyorsa rezerv nedenidir."))
    elif a["teslim"] in SIGORTA_TESLIM and istenen and not any(i["tur"] == "sigorta" for i in istenen):
        b.append(bulgu("Bilgi", "—", EVRAK_ADI["sigorta"], f"Teslim şekli {a['teslim']} ama akreditif sigorta belgesi istemiyor", a["teslim"], "", "Akreditif metnini kontrol edin."))
    # --- poliçe ve menşe
    pt = para(al(e.get("police", {}), "tutar"))
    if pt is not None and tutar is not None and pt != tutar:
        b.append(bulgu("Orta", "md. 14(d)", EVRAK_ADI["police"], "Poliçe tutarı fatura tutarından farklı", s(tutar), s(pt), "Akreditif farklı bir çekiş oranı öngörmüyorsa tutarlar aynı olmalıdır."))
    mu = al(e.get("mense", {}), "mense ulke", "mense")
    if a["mense"] and mu and not ayni(mu, a["mense"]):
        b.append(bulgu("Yüksek", "md. 14(d)", EVRAK_ADI["mense"], "Menşe ülke akreditiften farklı", a["mense"], metin(mu), "Menşe bilgisi akreditifle çelişemez."))

    # --- evraklar arası çelişki (md. 14(d))
    capraz = []
    for alan, ad, sayisal in (("koli", "Koli sayısı", True), ("brut agirlik", "Brüt ağırlık", True), ("net agirlik", "Net ağırlık", True), ("miktar", "Miktar", True),
                              ("alici", "Alıcı", False), ("gonderici", "Gönderici", False), ("konteyner no", "Konteyner no", False), ("gemi", "Gemi", False)):
        deg = {v["_ad"]: v[alan] for v in e.values() if alan in v}
        if len(deg) < 2:
            continue
        norm = {k: (para(v) if sayisal else katla(v)) for k, v in deg.items()}
        farkli = len(set(norm.values())) > 1
        if farkli and not sayisal:
            vals = list(norm.values())
            farkli = not all(x in y or y in x for x in vals for y in vals)
        capraz.append({"alan": ad, "degerler": deg, "farkli": farkli})
        if farkli:
            b.append(bulgu("Yüksek", "md. 14(d)", " / ".join(deg), f"{ad} evraklar arasında çelişiyor", "", "; ".join(f"{k}: {metin(v)}" for k, v in deg.items()),
                           "Evraklardaki bilgiler birbiriyle çelişemez; düzeltilmiş evrak alın."))
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    b.sort(key=lambda x: sira[x["onem"]])
    return {"akreditif": a, "ibraz": ibraz, "yukleme": yukleme, "son_ibraz": son_ibraz, "vade_fiili": vade_fiili, "ibraz_gun": ibraz_gun, "bulgular": b, "liste": liste,
            "capraz": capraz, "fatura_tutar": tutar, "evraklar": e}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

ONEM_ADI = {"Yüksek": "Rezerv riski", "Orta": "Kontrol edin", "Bilgi": "Bilgi"}


def rapor_yaz(cikti: Path, r: dict) -> None:
    a, wb = r["akreditif"], Workbook()
    oz = wb.active
    oz.title = "Özet"
    _baslik(oz, ["Bilgi", "Değer"], (38, 70))
    rezerv = sum(x["onem"] == "Yüksek" for x in r["bulgular"])
    kalan = (r["son_ibraz"] - r["ibraz"]).days
    for k, v in (("Akreditif no", a["no"]), ("Lehtar", a["lehtar"]), ("Amir", a["amir"]), ("Tutar", f"{s(a['tutar'])} {a['pb']}"), ("Fatura tutarı", s(r["fatura_tutar"])),
                 ("Son yükleme tarihi", g(a["son_yukleme"])), ("Yükleme tarihi (taşıma belgesi)", g(r["yukleme"])), ("Vade", g(a["vade"])),
                 ("İbraz süresi (gün)", r["ibraz_gun"]), ("En geç ibraz günü", g(r["son_ibraz"])), ("Planlanan ibraz", g(r["ibraz"])),
                 ("İbraz için kalan gün", kalan), ("Rezerv riski", rezerv), ("Kontrol edilecek", sum(x["onem"] == "Orta" for x in r["bulgular"]))):
        n = _satir(oz, [k, v])
        if k == "Rezerv riski":
            oz.cell(n, 2).fill = PatternFill("solid", fgColor=KIRMIZI if rezerv else YESIL)
        if k == "İbraz için kalan gün":
            oz.cell(n, 2).fill = PatternFill("solid", fgColor=KIRMIZI if kalan < 0 else SARI if kalan <= 3 else YESIL)
    oz.append([])
    oz.append(["Kod yalnız girilen alanları karşılaştırır; evrakın aslını okumaz. Rezerv kararını banka verir (inceleme süresi en çok 5 banka iş günü, UCP 600 md. 14(b))."])

    bu = wb.create_sheet("Rezervler")
    _baslik(bu, ["Durum", "UCP 600", "Evrak", "Bulgu", "Akreditifte", "Evrakta", "Öneri", "Düzeltildi"], (13, 16, 26, 44, 38, 38, 70, 10))
    for x in r["bulgular"]:
        _satir(bu, [ONEM_ADI[x["onem"]], x["madde"], x["evrak"], x["konu"], x["akreditifte"], x["evrakta"], x["oneri"], ""], dolgu={1: RENK[x["onem"]], 8: SARI}, sar=(4, 5, 6, 7))
    if not r["bulgular"]:
        bu.append(["Girilen alanlarda uyumsuzluk bulunmadı."])

    li = wb.create_sheet("Evrak Listesi")
    _baslik(li, ["İstenen Evrak", "İstenen Asıl", "İstenen Kopya", "İbraz Asıl", "İbraz Kopya", "Durum", "Akreditifteki Özel Şart (elle kontrol)"], (34, 9, 9, 9, 9, 13, 80))
    for x in r["liste"]:
        _satir(li, [x["ad"], x["asil"], x["kopya"], x["ibraz_asil"], x["ibraz_kopya"], x["durum"], x["sart"]], dolgu={6: YESIL if x["durum"] == "Tamam" else KIRMIZI}, sar=(7,))

    cp = wb.create_sheet("Çapraz Kontrol")
    adlar = [v["_ad"] for v in r["evraklar"].values()]
    _baslik(cp, ["Alan", *adlar, "Sonuç"], (18, *[26] * len(adlar), 12))
    for x in r["capraz"]:
        _satir(cp, [x["alan"], *[metin(x["degerler"].get(ad, "")) for ad in adlar], "Çelişki" if x["farkli"] else "Uyumlu"],
               dolgu={len(adlar) + 2: KIRMIZI if x["farkli"] else YESIL})
    kaydet(wb, cikti)


def calistir(a_yol: Path, i_yol: Path, e_yol: Path, ibraz: date, cikti: Path) -> dict:
    r = kontrol(akreditif_oku(a_yol), istenen_oku(i_yol), evraklar_oku(e_yol), ibraz)
    rapor_yaz(cikti, r)
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Akreditif şartlarını ibraz edilecek evraklarla karşılaştırır, olası rezervleri listeler (UCP 600).")
    p.add_argument("--akreditif", type=Path, default=ORNEK / "akreditif.csv", help="Alan;Değer: Akreditif No, Para Birimi, Tutar, Son Yükleme Tarihi, Vade, Lehtar, Amir, Mal Tanımı...")
    p.add_argument("--istenen", type=Path, default=ORNEK / "istenen_evraklar.csv", help="Evrak, Asıl, Kopya, Şart")
    p.add_argument("--evraklar", type=Path, default=ORNEK / "evraklar.csv", help="Evrak, Alan, Değer (hazırlanan evraklardaki bilgiler)")
    p.add_argument("--ibraz", help="Planlanan ibraz tarihi GG.AA.YYYY (varsayılan bugün; örnek veride 02.11.2026)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "akreditif_kontrol.xlsx")
    a = p.parse_args(argv)
    for y in (a.akreditif, a.istenen, a.evraklar):
        if not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    ibraz = tarih(a.ibraz) if a.ibraz else (date(2026, 11, 2) if a.akreditif == ORNEK / "akreditif.csv" else date.today())
    if not ibraz:
        print("[X] --ibraz GG.AA.YYYY biçiminde olmalı")
        return 2
    try:
        r = calistir(a.akreditif, a.istenen, a.evraklar, ibraz, a.cikti)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    rezerv = [x for x in r["bulgular"] if x["onem"] == "Yüksek"]
    print(f"[OK] {len(r['liste'])} istenen evrak · {len(rezerv)} rezerv riski · en geç ibraz {g(r['son_ibraz'])} (kalan {(r['son_ibraz'] - ibraz).days} gün)")
    for x in rezerv:
        print(f"[!] {x['evrak']} · {x['konu']} ({x['madde']})")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
