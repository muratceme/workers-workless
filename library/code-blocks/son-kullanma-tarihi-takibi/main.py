"""
Son Kullanma Tarihi Takibi — Gerisi Bende kod bloğu
Eczane ve Tıbbi Malzeme › Hastane Eczacısı
İlaç ve sarf stoklarını son kullanma tarihine göre listeler; yaklaşanları ve tüketim hızına göre süresi içinde tükenmeyecekleri işaretler.
  Kalan gün             = son kullanma tarihi − rapor tarihi (yalnız ay-yıl yazılmışsa ayın son günü)
  Günlük tüketim        = Günlük Tüketim sütunu · ya da Aylık Tüketim ÷ 30 · ya da --tuketim dosyasındaki dönem tüketimi ÷ --gun
  İlk giren ilk çıkar   = aynı malzemenin lotları son kullanma tarihine göre sırayla tüketilir (FEFO)
  Süresinde kullanılabilecek = en çok (günlük tüketim × kalan gün − kendisinden önce tüketilecek lotların miktarı), lot miktarını aşmadan
  Risk miktarı          = lot miktarı − süresinde kullanılabilecek miktar · Risk tutarı = risk miktarı × birim fiyat
  Durum                 = Süresi geçmiş · Tükenmeyecek (risk > 0) · Yaklaşıyor (kalan gün ≤ --yakin, 90) · Normal
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

STOK = {"kod": ("malzeme kodu", "ilac kodu", "stok kodu", "barkod", "kod"), "ad": ("malzeme adi", "ilac adi", "urun adi", "malzeme", "ad"), "lot": ("lot no", "lot", "seri no", "parti no", "lot seri no"),
        "skt": ("son kullanma tarihi", "skt", "miad", "son kullanma"), "miktar": ("miktar", "stok miktari", "eldeki miktar", "adet"), "fiyat": ("birim fiyat", "birim maliyet", "fiyat"),
        "depo": ("depo", "birim", "lokasyon"), "gunluk": ("gunluk tuketim", "gunluk ortalama tuketim"), "aylik": ("aylik tuketim", "aylik ortalama tuketim")}
TUKETIM = {"kod": ("malzeme kodu", "ilac kodu", "stok kodu", "barkod", "kod"), "miktar": ("tuketim", "tuketim miktari", "miktar", "cikis miktari")}
SIFIR = Decimal(0)


def skt_oku(x) -> date | None:
    """GG.AA.YYYY, YYYY-AA-GG, AA/YYYY, AA.YYYY ya da YYYY-AA; ay-yıl yazımında ayın son günü."""
    t = tarih(x)
    if t:
        return t
    s = metin(x)
    m = re.fullmatch(r"(\d{1,2})[./-](\d{4})", s) or None
    ay, yil = (int(m.group(1)), int(m.group(2))) if m else (None, None)
    if not m:
        m = re.fullmatch(r"(\d{4})[./-](\d{1,2})", s)
        ay, yil = (int(m.group(2)), int(m.group(1))) if m else (None, None)
    if not m or not 1 <= ay <= 12:
        return None
    return (date(yil + (ay == 12), ay % 12 + 1, 1) - timedelta(days=1))


def oku_stok(yol: Path) -> tuple[list[dict], list[dict]]:
    uy, l = [], []
    for r in kayitlar(yol, STOK, ("kod", "skt", "miktar")):
        kod = metin(r["kod"])
        if not kod:
            continue
        skt, miktar = skt_oku(r["skt"]), para(r["miktar"])
        d = {a: para(r[a]) if metin(r[a]) else None for a in ("fiyat", "gunluk", "aylik")}
        if skt is None or miktar is None or miktar < 0 or any(metin(r[a]) and (d[a] is None or d[a] < 0) for a in d):
            uy.append(uyari("Yüksek", "Stok satırı okunamadı", f"{kod} · {metin(r['lot'])}", f"Satır {r['_satir']}: son kullanma tarihi, miktar ya da fiyat / tüketim okunamadı. Satır alınmadı."))
            continue
        if miktar == 0:
            continue
        gunluk = d["gunluk"] if d["gunluk"] is not None else d["aylik"] / 30 if d["aylik"] is not None else None
        l.append({"kod": kod, "ad": metin(r["ad"]), "lot": metin(r["lot"]), "skt": skt, "miktar": miktar, "fiyat": d["fiyat"], "depo": metin(r["depo"]), "gunluk": gunluk})
    return l, uy


def oku_tuketim(yol: Path, gun: int) -> tuple[dict, list[dict]]:
    uy, d = [], defaultdict(lambda: SIFIR)
    for r in kayitlar(yol, TUKETIM, ("kod", "miktar")):
        kod = metin(r["kod"])
        if not kod:
            continue
        m = para(r["miktar"])
        if m is None or m < 0:
            uy.append(uyari("Yüksek", "Tüketim satırı okunamadı", kod, f"Satır {r['_satir']}: miktar okunamadı. Satır alınmadı."))
            continue
        d[katla(kod)] += m
    return {k: v / gun for k, v in d.items()}, uy


def takip(stoklar: list[dict], gun: date, yakin: int = 90, tuketim: dict | None = None) -> dict:
    if not stoklar:
        raise ValueError("Stok kaydı yok.")
    uy, gruplar = [], {}
    for s in stoklar:
        gruplar.setdefault(katla(s["kod"]), []).append(s)
    lotlar, malzemeler = [], []
    for k in sorted(gruplar):
        l = sorted(gruplar[k], key=lambda s: (s["skt"], s["lot"]))
        hiz = (tuketim or {}).get(k)
        if hiz is None:
            hizlar = [s["gunluk"] for s in l if s["gunluk"] is not None]
            hiz = max(hizlar) if hizlar else None
        ad = next((s["ad"] for s in l if s["ad"]), "")
        if hiz is None:
            uy.append(uyari("Bilgi", "Tüketim bilgisi yok", l[0]["kod"], f"{ad or l[0]['kod']}: günlük tüketim bilinmediği için tükenme riski hesaplanmadı; yalnız tarih kontrolü yapıldı."))
        once, toplam, risk, risk_tutar, risk_bilinmiyor = SIFIR, SIFIR, SIFIR, SIFIR, False
        for s in l:
            kalan = (s["skt"] - gun).days
            x = {**s, "ad": s["ad"] or ad, "kalan": kalan, "hiz": yuvarla(hiz) if hiz is not None else None, "kullan": None, "risk": None, "tutar": None}
            if kalan < 0:
                x["durum"], x["kullan"], x["risk"] = "Süresi geçmiş", SIFIR, s["miktar"]
                uy.append(uyari("Yüksek", "Süresi geçmiş stok", f"{s['kod']} · {s['lot']}", f"{x['ad']}: {s['miktar']:,.0f} birim, son kullanma {s['skt']:%d.%m.%Y}" + (f" ({s['depo']})" if s["depo"] else "") + ". Kullanımdan kaldırın."))
            elif hiz is not None:
                x["kullan"] = min(max(hiz * kalan - once, SIFIR), s["miktar"]).quantize(Decimal(1), rounding="ROUND_FLOOR")
                x["risk"] = s["miktar"] - x["kullan"]
                once += s["miktar"]
                x["durum"] = "Tükenmeyecek" if x["risk"] > 0 else "Yaklaşıyor" if kalan <= yakin else "Normal"
                if x["risk"] > 0:
                    uy.append(uyari("Orta", "Süresinde tükenmeyecek", f"{s['kod']} · {s['lot']}", f"{x['ad']}: {s['miktar']:,.0f} birimin tahmini {x['risk']:,.0f} birimi {s['skt']:%d.%m.%Y} tarihine kadar tüketilemez "
                                                                                                    f"(günlük tüketim {x['hiz']}). Başka birime transfer ya da tedarikçiye iade değerlendirin."))
            else:
                x["durum"] = "Yaklaşıyor" if kalan <= yakin else "Normal"
                risk_bilinmiyor = True
            if x["risk"] is not None and s["fiyat"] is not None:
                x["tutar"] = yuvarla(x["risk"] * s["fiyat"])
            toplam += s["miktar"]
            if x["risk"]:
                risk += x["risk"]
                risk_tutar += x["tutar"] or SIFIR
            lotlar.append(x)
        en_yakin = next((s for s in l if s["skt"] >= gun), None)
        malzemeler.append({"kod": l[0]["kod"], "ad": ad, "lot": len(l), "miktar": toplam, "hiz": yuvarla(hiz) if hiz is not None else None,
                           "stok_gun": yuvarla(sum((s["miktar"] for s in l if s["skt"] >= gun), SIFIR) / hiz) if hiz else None, "en_yakin": en_yakin["skt"] if en_yakin else None, "risk": risk, "risk_tutar": risk_tutar,
                           "risk_bilinmiyor": risk_bilinmiyor})
    lotlar.sort(key=lambda x: (x["skt"], katla(x["kod"]), x["lot"]))
    malzemeler.sort(key=lambda m: (-m["risk_tutar"], -m["risk"], katla(m["kod"])))
    ozet = {d: sum(1 for x in lotlar if x["durum"] == d) for d in ("Süresi geçmiş", "Tükenmeyecek", "Yaklaşıyor", "Normal")}
    return {"lotlar": lotlar, "malzemeler": malzemeler, "ozet": ozet, "risk_tutar": sum((m["risk_tutar"] for m in malzemeler), SIFIR), "uyarilar": uy}


DURUM_RENK = {"Süresi geçmiş": KIRMIZI, "Tükenmeyecek": SARI, "Yaklaşıyor": MAVI, "Normal": None}


def rapor_yaz(cikti: Path, r: dict, gun: date, yakin: int) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Lotlar"
    _baslik(ws, ["Son Kullanma", "Kalan Gün", "Durum", "Malzeme Kodu", "Malzeme Adı", "Lot / Seri", "Depo", "Miktar", "Günlük Tüketim", "Süresinde Kullanılabilecek", "Risk Miktarı", "Birim Fiyat", "Risk Tutarı"],
            (11, 8, 13, 13, 30, 12, 12, 9, 9, 11, 9, 10, 12), "D2")
    for x in r["lotlar"]:
        _satir(ws, [x["skt"], x["kalan"], x["durum"], x["kod"], x["ad"], x["lot"], x["depo"], x["miktar"], x["hiz"], x["kullan"], x["risk"], x["fiyat"], x["tutar"]], para_s=(9, 12, 13), tarih_s=(1,),
               dolgu={3: DURUM_RENK[x["durum"]]})
    ws.append([])
    ws.append([f"Rapor tarihi {gun:%d.%m.%Y}; yaklaşan: {yakin} gün ve altı. Lotlar son kullanma tarihine göre sırayla tüketilir (FEFO) varsayımıyla hesaplanmıştır."])
    ws = wb.create_sheet("Malzemeler")
    _baslik(ws, ["Malzeme Kodu", "Malzeme Adı", "Lot Sayısı", "Toplam Miktar", "Günlük Tüketim", "Stok Günü", "En Yakın SKT", "Risk Miktarı", "Risk Tutarı"], (13, 30, 8, 10, 9, 9, 11, 10, 12))
    for m in r["malzemeler"]:
        _satir(ws, [m["kod"], m["ad"], m["lot"], m["miktar"], m["hiz"], m["stok_gun"], m["en_yakin"], m["risk"] or None, m["risk_tutar"] or None], para_s=(5, 6, 9), tarih_s=(7,),
               dolgu={8: SARI if m["risk"] else None})
    uyari_sayfasi(wb, r["uyarilar"])
    kaydet(wb, cikti)


def calistir(stok_yol: Path, cikti: Path, gun: date, yakin: int = 90, tuketim_yol: Path | None = None, donem: int = 90) -> dict:
    stoklar, uy = oku_stok(stok_yol)
    tuketim, uy2 = oku_tuketim(tuketim_yol, donem) if tuketim_yol else (None, [])
    r = takip(stoklar, gun, yakin, tuketim)
    r["uyarilar"] = sirala(uy + uy2 + r["uyarilar"])
    rapor_yaz(cikti, r, gun, yakin)
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="İlaç ve sarf stoklarını son kullanma tarihine göre listeler; süresi içinde tükenmeyecekleri işaretler.")
    p.add_argument("--stok", type=Path, help="Malzeme Kodu, Malzeme Adı, Lot No, Son Kullanma Tarihi, Miktar, Birim Fiyat, Depo, Günlük / Aylık Tüketim")
    p.add_argument("--tuketim", type=Path, help="Dönem tüketimi (isteğe bağlı): Malzeme Kodu, Tüketim; günlük = toplam ÷ --gun")
    p.add_argument("--gun", type=int, default=90, help="Tüketim dosyasının kapsadığı gün sayısı (varsayılan 90)")
    p.add_argument("--tarih", help="Rapor tarihi GG.AA.YYYY (varsayılan bugün; örnek veride 30.09.2026)")
    p.add_argument("--yakin", type=int, default=90, help="Son kullanma tarihine bu kadar gün ve daha az kalanlar yaklaşan sayılır (varsayılan 90)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "son_kullanma_takibi.xlsx")
    a = p.parse_args(argv)
    yol = a.stok or ORNEK / "ilac_stok_lotlari.csv"
    gun = tarih(a.tarih) if a.tarih else (date(2026, 9, 30) if a.stok is None else date.today())
    if gun is None or a.gun < 1 or a.yakin < 0:
        print("[X] --tarih GG.AA.YYYY, --gun en az 1, --yakin sıfır ya da üstü olmalı")
        return 2
    for y in (yol, a.tuketim):
        if y and not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    try:
        r = calistir(yol, a.cikti, gun, a.yakin, a.tuketim, a.gun)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    o = r["ozet"]
    print(f"[OK] {len(r['malzemeler'])} malzeme · {len(r['lotlar'])} lot · süresi geçmiş {o['Süresi geçmiş']} · tükenmeyecek {o['Tükenmeyecek']} · yaklaşan {o['Yaklaşıyor']} · risk tutarı {r['risk_tutar']:,.2f}")
    for u in r["uyarilar"]:
        if u["onem"] in ("Yüksek", "Orta"):
            print(f"[!] {u['tur']}: {u['kim']} · {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
