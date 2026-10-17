"""
Ürün Ağacı (BOM) Karşılaştırma — Gerisi Bende kod bloğu
Ar-Ge ve Ürün Geliştirme › Ar-Ge Mühendisi

İki ürün ağacı sürümünü karşılaştırır; eklenen, çıkarılan ve miktarı, birimi, fiyatı ya da tanımı değişen kalemleri listeler.
  Eşleştirme anahtarı = üst kod + bileşen kodu   (üst kod sütunu yoksa yalnız bileşen kodu)
  Eklendi             = yalnız yeni sürümde;   Çıkarıldı = yalnız eski sürümde
  Birim değişti       = birim farklı (miktar farkı hesaplanmaz)
  Miktar değişti      = birim aynı, miktar farklı;   fark = yeni − eski;   fark % = fark ÷ eski × 100
  Fiyat değişti       = miktar ve birim aynı, birim maliyet farklı;   Tanım değişti = yalnız bileşen adı farklı
  Kalem maliyeti      = miktar × birim maliyet;   maliyet etkisi = yeni kalem maliyeti − eski kalem maliyeti
  Aynı anahtar bir dosyada birden çok satırdaysa miktarlar toplanır.
Kod çok seviyeli ağacı patlatmaz; her satır kendi üst koduyla karşılaştırılır. Kodlar ve birimler yazıldığı gibi eşleşir.
"""
from __future__ import annotations
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

SUTUN = {"ust": ("ust kod", "ust urun", "mamul kodu", "ana kod", "ust malzeme"), "kod": ("bilesen kodu", "malzeme kodu", "stok kodu", "parca kodu", "kod"), "ad": ("bilesen adi", "malzeme adi", "tanim", "ad"),
         "miktar": ("miktar", "birim miktar", "kullanim miktari"), "birim": ("birim", "olcu birimi"), "fiyat": ("birim maliyet", "birim fiyat", "maliyet")}
DURUM = ("Eklendi", "Çıkarıldı", "Birim değişti", "Miktar değişti", "Fiyat değişti", "Tanım değişti", "Aynı")
DURUM_RENK = {"Eklendi": YESIL, "Çıkarıldı": KIRMIZI, "Birim değişti": SARI, "Miktar değişti": SARI, "Fiyat değişti": MAVI, "Tanım değişti": MAVI, "Aynı": None}


def miktar_oku(x) -> Decimal | None:
    """0.250 → 0,25 · 1,5 → 1,5 · 1.250,5 → 1250,5 · 1.000 → 1000 (binlik); ondalık için virgül kullanın."""
    if isinstance(x, (int, float, Decimal)):
        return Decimal(str(x))
    s = metin(x).replace(" ", "")
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?[1-9]\d{0,2}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def oku(yol: Path, etiket: str) -> tuple[dict, list[dict]]:
    uy, d = [], {}
    for r in kayitlar(yol, SUTUN, ("kod", "miktar")):
        kod = metin(r["kod"])
        if not kod:
            continue
        kim = f"{metin(r['ust'])} › {kod}" if metin(r["ust"]) else kod
        miktar, fiyat = miktar_oku(r["miktar"]), para(r["fiyat"]) if metin(r["fiyat"]) else None
        if miktar is None or miktar <= 0 or (metin(r["fiyat"]) and (fiyat is None or fiyat < 0)):
            uy.append(uyari("Yüksek", "Satır okunamadı", kim, f"{etiket} ({yol.name}) satır {r['_satir']}: miktar '{metin(r['miktar'])}' ya da birim maliyet '{metin(r['fiyat'])}' okunamadı. Satır alınmadı."))
            continue
        k = (katla(r["ust"]), katla(kod))
        if k in d:
            d[k]["miktar"] += miktar
            uy.append(uyari("Bilgi", "Aynı kalem birden çok satırda", kim, f"{etiket} ({yol.name}) satır {r['_satir']}: miktar önceki satırla toplandı ({d[k]['miktar'].normalize():f})."))
        else:
            d[k] = {"ust": metin(r["ust"]), "kod": kod, "ad": metin(r["ad"]), "miktar": miktar, "birim": metin(r["birim"]), "fiyat": fiyat}
    return d, uy


def karsilastir(eski: dict, yeni: dict) -> dict:
    if not eski and not yeni:
        raise ValueError("İki ürün ağacı da boş.")
    satirlar = []
    for k in sorted(set(eski) | set(yeni)):
        e, y = eski.get(k), yeni.get(k)
        b = y or e
        s = {"ust": b["ust"], "kod": b["kod"], "ad": b["ad"], "eski_ad": e["ad"] if e else None, "eski_miktar": e["miktar"] if e else None, "yeni_miktar": y["miktar"] if y else None, "eski_birim": e["birim"] if e else None,
             "yeni_birim": y["birim"] if y else None, "fark": None, "fark_oran": None}
        ef, yf = (e or {}).get("fiyat"), (y or {}).get("fiyat")
        s["eski_fiyat"], s["yeni_fiyat"] = ef if ef is not None else (yf if e else None), yf if yf is not None else (ef if y else None)
        s["eski_tutar"] = e["miktar"] * s["eski_fiyat"] if e and s["eski_fiyat"] is not None else None
        s["yeni_tutar"] = y["miktar"] * s["yeni_fiyat"] if y and s["yeni_fiyat"] is not None else None
        s["etki"] = (s["yeni_tutar"] or 0) - (s["eski_tutar"] or 0) if s["eski_tutar"] is not None or s["yeni_tutar"] is not None else None
        if e is None:
            s["durum"] = "Eklendi"
        elif y is None:
            s["durum"] = "Çıkarıldı"
        elif katla(e["birim"]) != katla(y["birim"]):
            s["durum"] = "Birim değişti"
        elif e["miktar"] != y["miktar"]:
            s["durum"], s["fark"], s["fark_oran"] = "Miktar değişti", y["miktar"] - e["miktar"], yuvarla((y["miktar"] - e["miktar"]) * 100 / e["miktar"])
        elif ef is not None and yf is not None and ef != yf:
            s["durum"] = "Fiyat değişti"
        else:
            s["durum"] = "Tanım değişti" if katla(e["ad"]) != katla(y["ad"]) else "Aynı"
        satirlar.append(s)
    satirlar.sort(key=lambda s: (DURUM.index(s["durum"]), katla(s["ust"]), katla(s["kod"])))
    uy = [uyari("Orta", "Birim değişti", f"{s['ust']} › {s['kod']}" if s["ust"] else s["kod"], f"{s['ad']}: {s['eski_miktar'].normalize():f} {s['eski_birim']} → {s['yeni_miktar'].normalize():f} {s['yeni_birim']}. Miktarlar karşılaştırılmadı; "
                "birim dönüşümünü ve birim maliyeti kontrol edin.") for s in satirlar if s["durum"] == "Birim değişti"]
    fiyatli = any(s["etki"] is not None for s in satirlar)
    eski_t, yeni_t = sum((s["eski_tutar"] or 0 for s in satirlar), Decimal(0)), sum((s["yeni_tutar"] or 0 for s in satirlar), Decimal(0))
    ozet = {"eski": len(eski), "yeni": len(yeni), **{d: sum(1 for s in satirlar if s["durum"] == d) for d in DURUM}, "eski_tutar": eski_t if fiyatli else None, "yeni_tutar": yeni_t if fiyatli else None,
            "etki": yeni_t - eski_t if fiyatli else None, "etki_oran": yuvarla((yeni_t - eski_t) * 100 / eski_t) if fiyatli and eski_t else None,
            "fiyatsiz": sum(1 for s in satirlar if s["etki"] is None) if fiyatli else 0}
    if ozet["fiyatsiz"]:
        uy.append(uyari("Bilgi", "Maliyeti olmayan kalemler", "Toplam", f"{ozet['fiyatsiz']} kalemde iki sürümde de birim maliyet yok; maliyet etkisine girmedi."))
    return {"satirlar": satirlar, "degisenler": [s for s in satirlar if s["durum"] != "Aynı"], "ozet": ozet, "uyarilar": uy}


def rapor_yaz(cikti: Path, r: dict, eski_ad: str, yeni_ad: str) -> None:
    wb = Workbook()
    baslik = ["Durum", "Üst Kod", "Bileşen Kodu", "Bileşen Adı", "Eski Miktar", "Yeni Miktar", "Fark", "Fark %", "Eski Birim", "Yeni Birim", "Eski Birim Maliyet", "Yeni Birim Maliyet", "Eski Kalem Maliyeti", "Yeni Kalem Maliyeti",
              "Maliyet Etkisi", "Eski Tanım"]
    for i, (ad, alan) in enumerate((("Değişiklikler", "degisenler"), ("Tüm Kalemler", "satirlar"))):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = ad
        _baslik(ws, baslik, (16, 14, 16, 32, 11, 11, 10, 9, 9, 9, 13, 13, 14, 14, 14, 28))
        for s in r[alan]:
            _satir(ws, [s["durum"], s["ust"], s["kod"], s["ad"], s["eski_miktar"], s["yeni_miktar"], s["fark"], s["fark_oran"], s["eski_birim"], s["yeni_birim"], s["eski_fiyat"], s["yeni_fiyat"], s["eski_tutar"], s["yeni_tutar"], s["etki"],
                        s["eski_ad"] if s["durum"] == "Tanım değişti" or (s["eski_ad"] and katla(s["eski_ad"]) != katla(s["ad"])) else None], para_s=(8, 11, 12, 13, 14, 15), dolgu={1: DURUM_RENK[s["durum"]]})
    ws = wb.create_sheet("Özet")
    o = r["ozet"]
    _baslik(ws, ["Gösterge", "Değer"], (44, 18))
    for etiket, deger, para_mi in [(f"Eski sürüm kalem sayısı ({eski_ad})", o["eski"], False), (f"Yeni sürüm kalem sayısı ({yeni_ad})", o["yeni"], False)] + [(d, o[d], False) for d in DURUM] + \
            [("Eski sürüm toplam maliyeti", o["eski_tutar"], True), ("Yeni sürüm toplam maliyeti", o["yeni_tutar"], True), ("Maliyet etkisi", o["etki"], True), ("Maliyet etkisi %", o["etki_oran"], True)]:
        _satir(ws, [etiket, deger], para_s=(2,) if para_mi else (), dolgu={1: DURUM_RENK.get(etiket) if deger else None})
    uyari_sayfasi(wb, r["uyarilar"])
    kaydet(wb, cikti)


def calistir(eski_yol: Path, yeni_yol: Path, cikti: Path) -> dict:
    eski, uy1 = oku(eski_yol, "Eski sürüm")
    yeni, uy2 = oku(yeni_yol, "Yeni sürüm")
    r = karsilastir(eski, yeni)
    r["uyarilar"] = sirala(uy1 + uy2 + r["uyarilar"])
    rapor_yaz(cikti, r, eski_yol.name, yeni_yol.name)
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="İki ürün ağacı sürümünü karşılaştırıp eklenen, çıkarılan ve değişen kalemleri listeler.")
    p.add_argument("--eski", type=Path, help="Eski sürüm: Üst Kod, Bileşen Kodu, Bileşen Adı, Miktar, Birim, Birim Maliyet")
    p.add_argument("--yeni", type=Path, help="Yeni sürüm: aynı sütunlar")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "bom_karsilastirma.xlsx")
    a = p.parse_args(argv)
    if (a.eski is None) != (a.yeni is None):
        print("[X] --eski ve --yeni birlikte verilmeli")
        return 2
    eski_yol, yeni_yol = a.eski or ORNEK / "bom_v1.csv", a.yeni or ORNEK / "bom_v2.csv"
    for yol in (eski_yol, yeni_yol):
        if not yol.exists():
            print(f"[X] Dosya bulunamadı: {yol}")
            return 1
    try:
        r = calistir(eski_yol, yeni_yol, a.cikti)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    o = r["ozet"]
    print(f"[OK] Eski {o['eski']} kalem, yeni {o['yeni']} kalem · " + " · ".join(f"{d.lower()} {o[d]}" for d in DURUM) + (f" · maliyet etkisi {o['etki']:+,.2f}" if o["etki"] is not None else ""))
    for u in r["uyarilar"]:
        if u["onem"] in ("Yüksek", "Orta"):
            print(f"[!] {u['tur']}: {u['kim']} · {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
