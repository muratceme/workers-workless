"""
Hizmet Alımı Fatura Kontrolü — Workers / Workless kod bloğu
İdari İşler › İdari İşler Sorumlusu

Yemek, servis, temizlik ve güvenlik faturalarını sözleşme birim fiyatı ve fiilî kullanımla karşılaştırır:
  - Sözleşme kalemi: Hizmet + Kalem; fiyat geçerlilik aralığına göre (zam dönemleri ayrı satır).
  - Faturalama esası:
      Fiilî   → faturalanabilir miktar = kullanım toplamı; "Asgari Günlük" varsa her kullanım gününde en az o kadar
                (ör. asgari 120 öğün garantisi).
      Puantaj → kişi-ay = kişi-gün toplamı / "Ay Esası" (varsayılan 30); eksik gün kesintisi.
      Sabit   → ayda 1 birim.
  - Satır kontrolleri: birim fiyat ↔ sözleşme, miktar × fiyat = tutar, KDV oranı ↔ sözleşme, KDV tutarı, tevkifat
    tutarı (sözleşmedeki oran, ör. 9/10), mükerrer faturalama (aynı dönem ve kalem birden çok faturada).
  - Dönem mutabakatı: faturalanan − faturalanabilir miktar; fazlası sözleşme fiyatıyla itiraz tutarına çevrilir.
  - Faturası gelmemiş kullanım (tahakkuk gerekebilir) ve kullanım kaydı olmayan fatura.
Rapor: itiraz listesi (tedarikçiye gönderilebilir), dönem mutabakatı, fatura kontrolü, kullanım özeti, uyarılar.
İnternete bağlanmaz.

Kullanım:
    python main.py                                                  # örnek: Eylül 2026, 4 hizmet, 6 fatura
    python main.py --sozlesme sozlesme.xlsx --kullanim kullanim.xlsx --faturalar faturalar.xlsx --donem 2026-09
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
K2 = Decimal("0.01")
SIFIR = Decimal(0)
AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
AY_KATLI = {"ocak": 1, "subat": 2, "mart": 3, "nisan": 4, "mayis": 5, "haziran": 6, "temmuz": 7, "agustos": 8, "eylul": 9, "ekim": 10, "kasim": 11, "aralik": 12}

SOZLESME_SUTUNLARI = {"hizmet": ("hizmet",), "kalem": ("kalem", "hizmet kalemi"), "birim": ("birim", "olcu birimi"), "fiyat": ("birim fiyat", "fiyat"),
                      "kdv": ("kdv orani", "kdv"), "tevkifat": ("tevkifat", "tevkifat orani"), "bas": ("gecerlilik baslangic", "baslangic"),
                      "bit": ("gecerlilik bitis", "bitis"), "asgari": ("asgari gunluk", "asgari gunluk miktar", "garanti"),
                      "esas": ("faturalama esasi", "esas"), "ay_esasi": ("ay esasi", "ay esasi gun")}
KULLANIM_SUTUNLARI = {"tarih": ("tarih",), "hizmet": ("hizmet",), "kalem": ("kalem", "hizmet kalemi"), "miktar": ("miktar", "kisi gun", "adet"),
                      "kaynak": ("kaynak", "aciklama")}
FATURA_SUTUNLARI = {"no": ("fatura no", "belge no"), "tarih": ("fatura tarihi", "tarih"), "tedarikci": ("tedarikci", "satici", "firma"), "donem": ("donem", "hizmet donemi"),
                    "hizmet": ("hizmet",), "kalem": ("kalem", "hizmet kalemi"), "miktar": ("miktar",), "fiyat": ("birim fiyat", "fiyat"), "tutar": ("tutar", "matrah"),
                    "kdv_orani": ("kdv orani",), "kdv": ("kdv tutari", "kdv"), "tevkifat": ("tevkifat tutari", "tevkif edilen kdv", "tevkifat")}


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
    if x in (None, ""):
        return None
    if isinstance(x, (int, float)):
        return Decimal(str(x))
    s = str(x).strip().replace("TL", "").replace("%", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def oran(x) -> Decimal | None:
    """'%20', '20', '0,20' → 0.20; tevkifat '9/10' → 0.9."""
    s = str(x or "").strip()
    m = re.fullmatch(r"(\d+)\s*/\s*(\d+)", s)
    if m:
        return Decimal(m[1]) / Decimal(m[2])
    v = para(s)
    if v is None:
        return None
    return v / 100 if v > 1 else v


def tl(x: Decimal | None) -> str:
    return "—" if x is None else f"{x.quantize(K2, ROUND_HALF_UP):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def sayi(x: Decimal) -> str:
    """4.0000 → '4'; 5.9 → '5,9' (Türkçe ondalık)."""
    v = x.normalize()
    return f"{v:f}".replace(".", ",")


def yuzde_yaz(x: Decimal) -> str:
    return "%" + sayi(x * 100)


def kesir(x: Decimal) -> str:
    return sayi(x * 10) + "/10"


def metin(x) -> str:
    return str(x if x is not None else "").strip()


def donem_coz(x) -> tuple[int, int] | None:
    if isinstance(x, (date, datetime)):
        return x.year, x.month
    s = katla(x)
    m = re.fullmatch(r"(\d{4}) (\d{1,2})", s) or None
    if m:
        return int(m[1]), int(m[2])
    m = re.fullmatch(r"(\d{1,2}) (\d{4})", s)
    if m:
        return int(m[2]), int(m[1])
    m = re.fullmatch(r"([a-z]+) (\d{4})", s)
    if m and m[1] in AY_KATLI:
        return int(m[2]), AY_KATLI[m[1]]
    t = tarih(x)
    return (t.year, t.month) if t else None


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


def anahtar(hizmet, kalem) -> tuple[str, str]:
    return katla(hizmet), katla(kalem)


# ----------------------------------------------------------------------------
# Veri
# ----------------------------------------------------------------------------

@dataclass
class Fiyat:
    hizmet: str
    kalem: str
    birim: str
    fiyat: Decimal
    kdv: Decimal | None
    tevkifat: Decimal | None
    bas: date | None
    bit: date | None
    asgari: Decimal | None
    esas: str               # fiili | puantaj | sabit
    ay_esasi: Decimal


@dataclass
class FaturaSatiri:
    satir: int
    no: str
    tarih: date | None
    tedarikci: str
    donem: tuple[int, int] | None
    hizmet: str
    kalem: str
    miktar: Decimal
    fiyat: Decimal | None
    tutar: Decimal | None
    kdv_orani: Decimal | None
    kdv: Decimal | None
    tevkifat: Decimal | None
    sozlesme: Fiyat | None = None
    bulgular: list[dict] = field(default_factory=list)

    @property
    def itiraz(self) -> Decimal:
        return sum((b["tutar"] for b in self.bulgular if b["tutar"]), SIFIR)


def sozlesme_oku(yol: Path) -> list[Fiyat]:
    fiyatlar = []
    for r in kayitlar(yol, SOZLESME_SUTUNLARI, ("hizmet", "kalem", "fiyat")):
        f = para(r.get("fiyat"))
        if not metin(r.get("kalem")) or f is None:
            continue
        e = katla(r.get("esas"))
        esas = "puantaj" if "puantaj" in e or "kisi" in e else "sabit" if "sabit" in e or "goture" in e else "fiili"
        fiyatlar.append(Fiyat(metin(r["hizmet"]), metin(r["kalem"]), metin(r.get("birim")), f, oran(r.get("kdv")), oran(r.get("tevkifat")), tarih(r.get("bas")),
                              tarih(r.get("bit")), para(r.get("asgari")), esas, para(r.get("ay_esasi")) or Decimal(30)))
    return fiyatlar


def gecerli_fiyat(fiyatlar: list[Fiyat], hizmet: str, kalem: str, gun: date) -> Fiyat | None:
    a = anahtar(hizmet, kalem)
    adaylar = [f for f in fiyatlar if anahtar(f.hizmet, f.kalem) == a and (f.bas is None or f.bas <= gun) and (f.bit is None or gun <= f.bit)]
    return max(adaylar, key=lambda f: f.bas or date.min) if adaylar else None


def kullanim_oku(yol: Path | None) -> tuple[dict, list[dict]]:
    """(hizmet, kalem) → {tarih: miktar}"""
    k, uy = defaultdict(lambda: defaultdict(Decimal)), []
    if not yol:
        return k, uy
    for r in kayitlar(yol, KULLANIM_SUTUNLARI, ("tarih", "kalem", "miktar")):
        t, m = tarih(r.get("tarih")), para(r.get("miktar"))
        if not metin(r.get("kalem")):
            continue
        if not t or m is None:
            uy.append({"onem": "Orta", "tur": "Okunamayan kullanım", "kim": metin(r.get("kalem")), "aciklama": f"Satır {r['_satir']}: tarih veya miktar okunamadı"})
            continue
        k[anahtar(r.get("hizmet"), r.get("kalem"))][t] += m
    return k, uy


def faturalari_oku(yol: Path) -> tuple[list[FaturaSatiri], list[dict]]:
    satirlar, uy = [], []
    for r in kayitlar(yol, FATURA_SUTUNLARI, ("no", "kalem", "miktar")):
        if not metin(r.get("no")) or not metin(r.get("kalem")):
            continue
        m = para(r.get("miktar"))
        if m is None:
            uy.append({"onem": "Orta", "tur": "Okunamayan fatura satırı", "kim": metin(r["no"]), "aciklama": f"Satır {r['_satir']}: miktar okunamadı"})
            continue
        ft = tarih(r.get("tarih"))
        d = donem_coz(r.get("donem")) if metin(r.get("donem")) else ((ft.year, ft.month) if ft else None)
        satirlar.append(FaturaSatiri(r["_satir"], metin(r["no"]), ft, metin(r.get("tedarikci")), d, metin(r.get("hizmet")), metin(r.get("kalem")), m, para(r.get("fiyat")),
                                     para(r.get("tutar")), oran(r.get("kdv_orani")), para(r.get("kdv")), para(r.get("tevkifat"))))
    return satirlar, uy


# ----------------------------------------------------------------------------
# Kontrol
# ----------------------------------------------------------------------------

def faturalanabilir(f: Fiyat, gunluk: dict[date, Decimal]) -> tuple[Decimal, str]:
    if f.esas == "sabit":
        return Decimal(1), "Sabit aylık bedel"
    if f.esas == "puantaj":
        kg = sum(gunluk.values(), SIFIR)
        return (kg / f.ay_esasi).quantize(Decimal("0.0001"), ROUND_HALF_UP), f"{sayi(kg)} kişi-gün / {sayi(f.ay_esasi)}"
    if f.asgari:
        toplam = sum((max(m, f.asgari) for m in gunluk.values()), SIFIR)
        eksik = sum(1 for m in gunluk.values() if m < f.asgari)
        return toplam, f"{len(gunluk)} kullanım günü; {eksik} günde asgari {sayi(f.asgari)} uygulandı"
    return sum(gunluk.values(), SIFIR), f"{len(gunluk)} kullanım günü"


def kontrol_et(fiyatlar: list[Fiyat], kullanim: dict, faturalar: list[FaturaSatiri], donemler: set[tuple[int, int]] | None = None,
               tolerans: Decimal = Decimal("0.05")) -> dict:
    uy = []
    if donemler:
        faturalar = [x for x in faturalar if x.donem in donemler]
    for x in faturalar:
        if not x.donem:
            x.bulgular.append({"tur": "Dönem yok", "onem": "Orta", "tutar": None, "aciklama": "Hizmet dönemi okunamadı; sözleşme fiyatı ve kullanım eşleşmedi"})
            continue
        gun = date(x.donem[0], x.donem[1], 1)
        f = gecerli_fiyat(fiyatlar, x.hizmet, x.kalem, gun)
        x.sozlesme = f
        if not f:
            x.bulgular.append({"tur": "Sözleşmede yok", "onem": "Yüksek", "tutar": x.tutar,
                               "aciklama": f"'{x.hizmet} / {x.kalem}' için {x.donem[1]:02d}.{x.donem[0]} döneminde geçerli sözleşme fiyatı yok"})
            continue
        if x.fiyat is not None and abs(x.fiyat - f.fiyat) > K2 / 2:
            fark = ((x.fiyat - f.fiyat) * x.miktar).quantize(K2, ROUND_HALF_UP)
            x.bulgular.append({"tur": "Fiyat farkı", "onem": "Yüksek" if fark > 0 else "Bilgi", "tutar": max(fark, SIFIR),
                               "aciklama": f"Fatura birim fiyatı {tl(x.fiyat)}, sözleşme {tl(f.fiyat)} → {tl(fark)} TL"})
        if x.fiyat is not None and x.tutar is not None and abs(x.miktar * x.fiyat - x.tutar) > tolerans:
            x.bulgular.append({"tur": "Hesap hatası", "onem": "Orta", "tutar": max(SIFIR, (x.tutar - x.miktar * x.fiyat).quantize(K2, ROUND_HALF_UP)),
                               "aciklama": f"{sayi(x.miktar)} × {tl(x.fiyat)} = {tl(x.miktar * x.fiyat)}; faturada {tl(x.tutar)}"})
        if x.kdv_orani is not None and f.kdv is not None and x.kdv_orani != f.kdv:
            x.bulgular.append({"tur": "KDV oranı", "onem": "Yüksek", "tutar": None,
                               "aciklama": f"Faturada {yuzde_yaz(x.kdv_orani)}, sözleşmede {yuzde_yaz(f.kdv)}; faturanın düzeltilmesini isteyin"})
        kdv_oran = x.kdv_orani if x.kdv_orani is not None else f.kdv
        if x.kdv is not None and x.tutar is not None and kdv_oran is not None and abs(x.tutar * kdv_oran - x.kdv) > tolerans:
            x.bulgular.append({"tur": "KDV tutarı", "onem": "Orta", "tutar": None,
                               "aciklama": f"{tl(x.tutar)} × {yuzde_yaz(kdv_oran)} = {tl(x.tutar * kdv_oran)}; faturada {tl(x.kdv)}"})
        if f.tevkifat and x.kdv is not None:
            beklenen = (x.kdv * f.tevkifat).quantize(K2, ROUND_HALF_UP)
            if x.tevkifat is None:
                x.bulgular.append({"tur": "Tevkifat yok", "onem": "Orta", "tutar": None,
                                   "aciklama": f"Sözleşmede tevkifat {kesir(f.tevkifat)}; faturada tevkifat tutarı yok (tutar sınırının altındaysa uygulanmaz)"})
            elif abs(beklenen - x.tevkifat) > tolerans:
                x.bulgular.append({"tur": "Tevkifat tutarı", "onem": "Yüksek", "tutar": None,
                                   "aciklama": f"Beklenen tevkifat {tl(x.kdv)} × {kesir(f.tevkifat)} = {tl(beklenen)}; faturada {tl(x.tevkifat)}"})

    # Mükerrer: aynı dönem + kalem birden çok faturada; aynı faturada aynı kalem iki kez
    gruplar = defaultdict(list)
    for x in faturalar:
        if x.donem:
            gruplar[(x.donem, anahtar(x.hizmet, x.kalem))].append(x)
    for lst in gruplar.values():
        faturano = sorted({x.no for x in lst})
        if len(faturano) > 1:
            for x in lst[1:]:
                if x.no != lst[0].no:
                    x.bulgular.append({"tur": "Mükerrer faturalama", "onem": "Yüksek", "tutar": None,
                                       "aciklama": f"Aynı dönem ve kalem {', '.join(faturano)} faturalarında; miktar farkı mutabakatta değerlendirildi"})

    # Dönem mutabakatı
    mutabakat = []
    for (donem, (hz, kl)), lst in sorted(gruplar.items()):
        f = lst[0].sozlesme
        if not f:
            continue
        gunluk = {t: m for t, m in kullanim.get((hz, kl), {}).items() if (t.year, t.month) == donem}
        fat_miktar = sum((x.miktar for x in lst), SIFIR)
        hak, aciklama = faturalanabilir(f, gunluk)
        fark = fat_miktar - hak
        satir = {"donem": donem, "hizmet": f.hizmet, "kalem": f.kalem, "birim": f.birim, "esas": f.esas, "fiili": sum(gunluk.values(), SIFIR), "hak": hak,
                 "faturalanan": fat_miktar, "fark": fark, "fiyat": f.fiyat, "itiraz": SIFIR, "aciklama": aciklama, "faturalar": sorted({x.no for x in lst})}
        if not gunluk and f.esas != "sabit":
            satir["aciklama"] = "Kullanım kaydı yok"
            uy.append({"onem": "Yüksek", "tur": "Kullanım kaydı yok", "kim": f"{f.hizmet} / {f.kalem}",
                       "aciklama": f"{donem[1]:02d}.{donem[0]}: {sayi(fat_miktar)} {f.birim} faturalanmış, kullanım kaydı yok; puantaj / yoklama isteyin"})
        elif fark > Decimal("0.0001"):
            satir["itiraz"] = (fark * f.fiyat).quantize(K2, ROUND_HALF_UP)
            hedef = next((x for x in reversed(lst) if any(b["tur"] == "Mükerrer faturalama" for b in x.bulgular)), None) or max(lst, key=lambda x: x.miktar)
            hedef.bulgular.append({"tur": "Fazla miktar", "onem": "Yüksek", "tutar": satir["itiraz"],
                                   "aciklama": f"Faturalanan {sayi(fat_miktar)} {f.birim}, faturalanabilir {sayi(hak)} ({aciklama}); fark {sayi(fark)} × {tl(f.fiyat)}"})
        elif fark < 0:
            uy.append({"onem": "Bilgi", "tur": "Eksik faturalama", "kim": f"{f.hizmet} / {f.kalem}",
                       "aciklama": f"{donem[1]:02d}.{donem[0]}: faturalanabilir {sayi(hak)}, faturalanan {sayi(fat_miktar)}; sonraki faturada gelebilir (tahakkuk)"})
        mutabakat.append(satir)
    fat_kalemleri = {(x.donem, anahtar(x.hizmet, x.kalem)) for x in faturalar}
    kul_donemler = donemler or {x.donem for x in faturalar if x.donem}
    for (hz, kl), gunluk in kullanim.items():
        for d in sorted({(t.year, t.month) for t in gunluk} & kul_donemler):
            if (d, (hz, kl)) not in fat_kalemleri:
                f = next((f for f in fiyatlar if anahtar(f.hizmet, f.kalem) == (hz, kl)), None)
                ad = f"{f.hizmet} / {f.kalem}" if f else f"{hz} / {kl}"
                miktar = sum((m for t, m in gunluk.items() if (t.year, t.month) == d), SIFIR)
                uy.append({"onem": "Bilgi", "tur": "Faturası gelmemiş kullanım", "kim": ad,
                           "aciklama": f"{d[1]:02d}.{d[0]}: {sayi(miktar)} kullanım var, fatura yok; gider tahakkuku gerekebilir"})
    for x in faturalar:
        for b in x.bulgular:
            uy.append({"onem": b["onem"], "tur": b["tur"], "kim": f"{x.no} · {x.kalem}", "aciklama": b["aciklama"]})
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    uy.sort(key=lambda u: (sira[u["onem"]], u["tur"], u["kim"]))
    return {"faturalar": faturalar, "mutabakat": mutabakat, "kullanim": kullanim, "uyarilar": uy, "fiyatlar": fiyatlar}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
RENK = {"Yüksek": "FDE2E1", "Orta": "FFF4CE", "Bilgi": "E8F0FE"}
UST = Alignment(vertical="top", wrap_text=True)
PF, MF = "#,##0.00", "#,##0.####"
ESAS = {"fiili": "Fiilî", "puantaj": "Puantaj", "sabit": "Sabit"}


def _baslik(ws, basliklar, genislik=()):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def _f(x):
    return None if x is None else float(x)


def rapor_yaz(cikti: Path, s: dict) -> None:
    wb = Workbook()
    it = wb.active
    it.title = "İtiraz Listesi"
    _baslik(it, ["Tedarikçi", "Fatura No", "Fatura Tarihi", "Dönem", "Hizmet", "Kalem", "Bulgu", "İtiraz Tutarı (KDV hariç)", "Açıklama", "Tedarikçi Cevabı"],
            (24, 14, 11, 9, 11, 26, 20, 14, 80, 22))
    toplam = SIFIR
    for x in sorted(s["faturalar"], key=lambda x: (katla(x.tedarikci), x.no, x.satir)):
        for b in x.bulgular:
            if b["onem"] == "Bilgi":
                continue
            it.append([x.tedarikci, x.no, x.tarih, f"{x.donem[1]:02d}.{x.donem[0]}" if x.donem else "", x.hizmet, x.kalem, b["tur"], _f(b["tutar"]), b["aciklama"], ""])
            it.cell(it.max_row, 3).number_format = "DD.MM.YYYY"
            it.cell(it.max_row, 8).number_format = PF
            it.cell(it.max_row, 7).fill = PatternFill("solid", fgColor=RENK[b["onem"]])
            it.cell(it.max_row, 9).alignment = UST
            it.cell(it.max_row, 10).fill = PatternFill("solid", fgColor="FFF4CE")
            toplam += b["tutar"] or SIFIR
    it.append([])
    it.append(["Toplam itiraz (KDV hariç)", None, None, None, None, None, None, float(toplam)])
    it.cell(it.max_row, 1).font = it.cell(it.max_row, 8).font = Font(bold=True)
    it.cell(it.max_row, 8).number_format = PF

    mu = wb.create_sheet("Dönem Mutabakatı")
    _baslik(mu, ["Dönem", "Hizmet", "Kalem", "Birim", "Esas", "Fiilî Kullanım", "Faturalanabilir", "Faturalanan", "Fark", "Sözleşme Fiyatı", "İtiraz Tutarı",
                 "Hesap", "Faturalar"], (9, 11, 26, 9, 8, 11, 12, 11, 9, 12, 13, 40, 22))
    for m in s["mutabakat"]:
        mu.append([f"{m['donem'][1]:02d}.{m['donem'][0]}", m["hizmet"], m["kalem"], m["birim"], ESAS[m["esas"]], float(m["fiili"]), float(m["hak"]),
                   float(m["faturalanan"]), float(m["fark"]), float(m["fiyat"]), float(m["itiraz"]), m["aciklama"], ", ".join(m["faturalar"])])
        for j in (6, 7, 8, 9):
            mu.cell(mu.max_row, j).number_format = MF
        mu.cell(mu.max_row, 10).number_format = mu.cell(mu.max_row, 11).number_format = PF
        if m["itiraz"] > 0:
            mu.cell(mu.max_row, 9).fill = PatternFill("solid", fgColor="FDE2E1")

    fk = wb.create_sheet("Fatura Kontrolü")
    _baslik(fk, ["Fatura No", "Tarih", "Tedarikçi", "Dönem", "Hizmet", "Kalem", "Miktar", "Birim Fiyat", "Sözleşme Fiyatı", "Tutar", "KDV %", "KDV", "Tevkifat",
                 "Beklenen Tevkifat", "Bulgular"], (14, 11, 22, 9, 11, 26, 9, 11, 11, 13, 6, 12, 12, 12, 60))
    for x in sorted(s["faturalar"], key=lambda x: (x.no, x.satir)):
        f = x.sozlesme
        bek = (x.kdv * f.tevkifat).quantize(K2, ROUND_HALF_UP) if f and f.tevkifat and x.kdv is not None else None
        fk.append([x.no, x.tarih, x.tedarikci, f"{x.donem[1]:02d}.{x.donem[0]}" if x.donem else "", x.hizmet, x.kalem, float(x.miktar), _f(x.fiyat), _f(f.fiyat if f else None),
                   _f(x.tutar), _f(x.kdv_orani * 100 if x.kdv_orani is not None else None), _f(x.kdv), _f(x.tevkifat), _f(bek),
                   "; ".join(b["tur"] for b in x.bulgular) or "Uygun"])
        fk.cell(fk.max_row, 2).number_format = "DD.MM.YYYY"
        for j in (8, 9, 10, 12, 13, 14):
            fk.cell(fk.max_row, j).number_format = PF
        fk.cell(fk.max_row, 15).fill = PatternFill("solid", fgColor="E3F4E1" if not x.bulgular else "FFF4CE")
    fk.auto_filter.ref = f"A1:O{fk.max_row}"

    ko = wb.create_sheet("Kullanım Özeti")
    donemler = sorted({(t.year, t.month) for g in s["kullanim"].values() for t in g})
    _baslik(ko, ["Hizmet", "Kalem"] + [f"{AYLAR[m - 1][:3]} {y}" for y, m in donemler], [12, 28] + [11] * len(donemler))
    adlar = {anahtar(f.hizmet, f.kalem): (f.hizmet, f.kalem) for f in s["fiyatlar"]}
    for k, g in sorted(s["kullanim"].items()):
        h, kl = adlar.get(k, k)
        ko.append([h, kl] + [float(sum((m for t, m in g.items() if (t.year, t.month) == d), SIFIR)) for d in donemler])

    uy = wb.create_sheet("Uyarılar")
    _baslik(uy, ["Önem", "Tür", "Fatura / Kalem", "Açıklama"], (9, 24, 34, 100))
    for u in s["uyarilar"]:
        uy.append([u["onem"], u["tur"], u["kim"], u["aciklama"]])
        uy.cell(uy.max_row, 1).fill = PatternFill("solid", fgColor=RENK[u["onem"]])
        uy.cell(uy.max_row, 4).alignment = UST
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(sozlesme_yolu: Path, fatura_yolu: Path, cikti: Path, kullanim_yolu: Path | None = None, donemler: set[tuple[int, int]] | None = None) -> dict:
    fiyatlar = sozlesme_oku(sozlesme_yolu)
    if not fiyatlar:
        raise ValueError(f"{sozlesme_yolu.name}: sözleşme kalemi bulunamadı")
    faturalar, uy1 = faturalari_oku(fatura_yolu)
    if not faturalar:
        raise ValueError(f"{fatura_yolu.name}: fatura satırı bulunamadı")
    kullanim, uy2 = kullanim_oku(kullanim_yolu)
    s = kontrol_et(fiyatlar, kullanim, faturalar, donemler)
    s["uyarilar"] = uy1 + uy2 + s["uyarilar"]
    rapor_yaz(cikti, s)
    return s


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Hizmet alımı faturalarını sözleşme fiyatı ve fiilî kullanımla karşılaştırır.")
    p.add_argument("--sozlesme", type=Path, default=ORNEK / "sozlesme.csv",
                   help="Hizmet, Kalem, Birim, Birim Fiyat, KDV Oranı, Tevkifat (ör. 9/10), Geçerlilik Başlangıç, Geçerlilik Bitiş, Asgari Günlük, "
                        "Faturalama Esası (Fiilî / Puantaj / Sabit), Ay Esası")
    p.add_argument("--faturalar", type=Path, default=ORNEK / "faturalar.csv",
                   help="Fatura No, Fatura Tarihi, Tedarikçi, Dönem, Hizmet, Kalem, Miktar, Birim Fiyat, Tutar, KDV Oranı, KDV Tutarı, Tevkifat Tutarı")
    p.add_argument("--kullanim", type=Path, help="Tarih, Hizmet, Kalem, Miktar (öğün, sefer, saat veya puantajda kişi-gün), Kaynak")
    p.add_argument("--donem", nargs="*", default=[], help="Yalnız bu dönemler (ör. 2026-09 08.2026)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "hizmet_fatura_kontrolu.xlsx")
    a = p.parse_args(argv)
    ornek = a.sozlesme == ORNEK / "sozlesme.csv"
    kul = a.kullanim or (ORNEK / "kullanim.csv" if ornek else None)
    for y in (a.sozlesme, a.faturalar, kul):
        if y and not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    donemler = set()
    for d in a.donem:
        c = donem_coz(d)
        if not c:
            print(f"[X] Dönem okunamadı: {d} (ör. 2026-09)")
            return 2
        donemler.add(c)
    try:
        s = calistir(a.sozlesme, a.faturalar, a.cikti, kul, donemler or None)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    itiraz = sum((x.itiraz for x in s["faturalar"]), SIFIR)
    print(f"[OK] {len({x.no for x in s['faturalar']})} fatura, {len(s['faturalar'])} satır · bulgulu satır {sum(bool(x.bulgular) for x in s['faturalar'])} · "
          f"itiraz tutarı {tl(itiraz)} TL (KDV hariç)")
    for u in s["uyarilar"]:
        if u["onem"] == "Yüksek":
            print(f"[!] {u['kim']} · {u['tur']}: {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
