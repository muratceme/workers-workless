"""
Abonelik ve Sözleşme Takibi — Gerisi Bende kod bloğu
İdari İşler › İdari İşler Sorumlusu
Kira, elektrik, su, internet, bakım ve benzeri sözleşmelerin tutar, vade, fesih ihbarı ve yenileme tarihlerini izler.
  İhbar son günü = bitiş tarihi − fesih ihbar süresi (gün); yalnız otomatik yenilenen sözleşmelerde: bu güne kadar bildirim yapılmazsa yenilenir
  Durum          = Süresi doldu · İhbar süresi geçti (otomatik yenilenecek) · İhbar yaklaşıyor (≤ --ihbar 30 gün) · Bitiş yaklaşıyor (≤ --bitis 60 gün) · Devam ediyor · Süresiz
  Aylık maliyet  = aylık tutar (yıllık / çeyreklik yazılmışsa aya çevrilir); kategori ve para birimi bazında toplam
  Artış tarihi   = başlangıcın yıl dönümü; artış esası (TÜFE / sabit %) yazılıysa listede gösterilir
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
SUTUN = {"ad": ("sozlesme", "abonelik", "hizmet", "ad"), "kategori": ("kategori", "tur", "gider turu"), "tedarikci": ("tedarikci", "firma", "karsi taraf", "saglayici"),
         "abone": ("abone no", "sozlesme no", "tesisat no"), "bas": ("baslangic tarihi", "baslangic"), "bitis": ("bitis tarihi", "bitis", "sona erme"),
         "tutar": ("tutar", "aylik tutar", "bedel"), "periyot": ("odeme periyodu", "periyot", "siklik"), "pb": ("para birimi", "doviz", "pb"),
         "oto": ("otomatik yenileme", "oto yenileme", "yenileme"), "ihbar": ("fesih ihbar suresi", "ihbar suresi", "ihbar gun"), "artis": ("artis esasi", "artis", "fiyat artisi"),
         "sorumlu": ("sorumlu",), "lokasyon": ("lokasyon", "sube", "adres")}
CARPAN = {"aylik": Decimal(1), "ay": Decimal(1), "yillik": Decimal(1) / 12, "yil": Decimal(1) / 12, "ceyrek": Decimal(1) / 3, "3 aylik": Decimal(1) / 3, "6 aylik": Decimal(1) / 6}
SIFIR = Decimal(0)


def oku(yol: Path) -> tuple[list[dict], list[dict]]:
    uy, l = [], []
    for r in kayitlar(yol, SUTUN, ("ad", "tutar")):
        ad = metin(r["ad"])
        if not ad:
            continue
        bas, bit = (tarih(r[a]) if metin(r[a]) else None for a in ("bas", "bitis"))
        tutar, ihbar = para(r["tutar"]), (para(r["ihbar"]) if metin(r["ihbar"]) else SIFIR)
        per = katla(r["periyot"]) or "aylik"
        if tutar is None or tutar < 0 or ihbar is None or ihbar < 0 or per not in CARPAN or (metin(r["bas"]) and bas is None) or (metin(r["bitis"]) and bit is None):
            uy.append(uyari("Yüksek", "Satır okunamadı", ad, f"Satır {r['_satir']}: tutar, periyot ('{metin(r['periyot'])}'), ihbar süresi ya da tarihlerden biri okunamadı."))
            continue
        if bas and bit and bit < bas:
            uy.append(uyari("Yüksek", "Bitiş başlangıçtan önce", ad, f"Başlangıç {bas:%d.%m.%Y}, bitiş {bit:%d.%m.%Y}."))
            continue
        l.append({"ad": ad, "kategori": metin(r["kategori"]) or "Diğer", "tedarikci": metin(r["tedarikci"]), "abone": metin(r["abone"]), "bas": bas, "bitis": bit, "tutar": tutar,
                  "aylik": yuvarla(tutar * CARPAN[per]), "pb": metin(r["pb"]).upper() or "TRY", "oto": katla(r["oto"]) in ("evet", "e", "var", "x", "otomatik"), "ihbar": int(ihbar),
                  "artis": metin(r["artis"]), "sorumlu": metin(r["sorumlu"]), "lokasyon": metin(r["lokasyon"])})
    return l, uy


def yildonumu(bas: date, bugun: date) -> date:
    y = bugun.year
    while True:
        try:
            g = bas.replace(year=y)
        except ValueError:
            g = date(y, 3, 1)
        if g >= bugun:
            return g
        y += 1


def analiz(sozlesmeler: list[dict], bugun: date, ihbar_esik: int = 30, bitis_esik: int = 60) -> dict:
    if not sozlesmeler:
        raise ValueError("Okunabilen sözleşme yok.")
    uy = []
    for s in sozlesmeler:
        kim = f"{s['ad']}" + (f" ({s['tedarikci']})" if s["tedarikci"] else "")
        s["ihbar_son"] = s["bitis"] - timedelta(days=s["ihbar"]) if s["bitis"] and s["ihbar"] and s["oto"] else None
        s["kalan"] = (s["bitis"] - bugun).days if s["bitis"] else None
        s["ihbar_kalan"] = (s["ihbar_son"] - bugun).days if s["ihbar_son"] else None
        s["artis_tarihi"] = yildonumu(s["bas"], bugun) if s["bas"] and s["artis"] else None
        if s["bitis"] is None:
            s["durum"] = "Süresiz"
        elif s["kalan"] < 0:
            s["durum"] = "Süresi doldu"
        elif s["ihbar_kalan"] is not None and s["ihbar_kalan"] < 0:
            s["durum"] = "İhbar süresi geçti"
        elif s["ihbar_kalan"] is not None and s["ihbar_kalan"] <= ihbar_esik:
            s["durum"] = "İhbar yaklaşıyor"
        elif s["kalan"] <= bitis_esik:
            s["durum"] = "Bitiş yaklaşıyor"
        else:
            s["durum"] = "Devam ediyor"
        if s["durum"] == "Süresi doldu":
            uy.append(uyari("Yüksek" if not s["oto"] else "Orta", "Süresi doldu", kim, f"Bitiş {s['bitis']:%d.%m.%Y}." + (" Otomatik yenileme işaretli; yeni dönem bitiş tarihini güncelleyin." if s["oto"] else " Hizmet kesintisi riski; yenileme / yeni sözleşme durumunu kontrol edin.")))
        elif s["durum"] == "İhbar yaklaşıyor":
            uy.append(uyari("Orta", "Fesih ihbarı son günü yaklaşıyor", kim, f"Değiştirmek ya da feshetmek için son bildirim günü {s['ihbar_son']:%d.%m.%Y} ({s['ihbar_kalan']} gün); bitiş {s['bitis']:%d.%m.%Y}."))
        elif s["durum"] == "İhbar süresi geçti":
            uy.append(uyari("Bilgi", "İhbar süresi geçti", kim, f"Son bildirim günü {s['ihbar_son']:%d.%m.%Y} geçti; sözleşme {s['bitis']:%d.%m.%Y} sonrasında koşullarına göre " + ("yenilenecek." if s["oto"] else "sona erecek ya da yenilenecek.")))
        elif s["durum"] == "Bitiş yaklaşıyor":
            uy.append(uyari("Orta", "Bitiş yaklaşıyor", kim, f"Bitiş {s['bitis']:%d.%m.%Y} ({s['kalan']} gün)" + ("; otomatik yenilenecek." if s["oto"] else "; yenileme gerekli.")))
        if s["artis_tarihi"] and (s["artis_tarihi"] - bugun).days <= ihbar_esik:
            uy.append(uyari("Bilgi", "Fiyat artış dönemi", kim, f"Yıl dönümü {s['artis_tarihi']:%d.%m.%Y}; artış esası: {s['artis']}. Yeni tutarı hesaplayıp bütçeye yansıtın."))
    kat = {}
    for s in sozlesmeler:
        if s["durum"] == "Süresi doldu" and not s["oto"]:
            continue
        k = kat.setdefault((s["kategori"], s["pb"]), {"kategori": s["kategori"], "pb": s["pb"], "adet": 0, "aylik": SIFIR})
        k["adet"] += 1
        k["aylik"] += s["aylik"]
    for k in kat.values():
        k["yillik"] = k["aylik"] * 12
    sira = {"Süresi doldu": 0, "İhbar yaklaşıyor": 1, "Bitiş yaklaşıyor": 2, "İhbar süresi geçti": 3, "Devam ediyor": 4, "Süresiz": 5}
    return {"sozlesmeler": sorted(sozlesmeler, key=lambda s: (sira[s["durum"]], s["ihbar_son"] or s["bitis"] or date.max, katla(s["ad"]))),
            "kategoriler": sorted(kat.values(), key=lambda k: (k["pb"], -k["aylik"])), "ozet": {d: sum(1 for s in sozlesmeler if s["durum"] == d) for d in sira}, "uyarilar": uy}


def rapor_yaz(cikti: Path, r: dict) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sözleşmeler"
    _baslik(ws, ["Sözleşme", "Kategori", "Tedarikçi", "Abone / Sözleşme No", "Lokasyon", "Başlangıç", "Bitiş", "Kalan Gün", "İhbar Süresi", "İhbar Son Günü", "Durum", "Otomatik Yenileme",
                 "Tutar", "Aylık Karşılık", "Para Birimi", "Artış Esası", "Artış Tarihi", "Sorumlu"], (26, 12, 18, 14, 12, 11, 11, 8, 8, 11, 17, 9, 12, 12, 7, 16, 11, 12), "B2")
    renk = {"Süresi doldu": KIRMIZI, "İhbar yaklaşıyor": SARI, "Bitiş yaklaşıyor": SARI, "İhbar süresi geçti": MAVI}
    for s in r["sozlesmeler"]:
        _satir(ws, [s["ad"], s["kategori"], s["tedarikci"], s["abone"], s["lokasyon"], s["bas"], s["bitis"], s["kalan"], s["ihbar"] or None, s["ihbar_son"], s["durum"],
                    "Evet" if s["oto"] else "Hayır", s["tutar"], s["aylik"], s["pb"], s["artis"], s["artis_tarihi"], s["sorumlu"]], para_s=(13, 14), tarih_s=(6, 7, 10, 17),
               dolgu={11: renk.get(s["durum"])})
    ws = wb.create_sheet("Maliyet")
    _baslik(ws, ["Kategori", "Para Birimi", "Sözleşme", "Aylık", "Yıllık"], (16, 8, 9, 14, 15))
    for k in r["kategoriler"]:
        _satir(ws, [k["kategori"], k["pb"], k["adet"], k["aylik"], k["yillik"]], para_s=(4, 5))
    ws.append([])
    ws.append(["Süresi dolmuş ve otomatik yenilenmeyen sözleşmeler maliyete dahil edilmemiştir."])
    uyari_sayfasi(wb, r["uyarilar"])
    kaydet(wb, cikti)


def calistir(yol: Path, cikti: Path, bugun: date, ihbar: int = 30, bitis: int = 60) -> dict:
    s, uy = oku(yol)
    r = analiz(s, bugun, ihbar, bitis)
    r["uyarilar"] = sirala(uy + r["uyarilar"])
    rapor_yaz(cikti, r)
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Sözleşme ve aboneliklerin bitiş, fesih ihbarı ve maliyetlerini izler.")
    p.add_argument("--sozlesmeler", type=Path, help="Sözleşme, Kategori, Tedarikçi, Abone No, Lokasyon, Başlangıç, Bitiş, Tutar, Ödeme Periyodu (Aylık / Yıllık / 3 aylık), "
                                                    "Para Birimi, Otomatik Yenileme, Fesih İhbar Süresi (gün), Artış Esası, Sorumlu")
    p.add_argument("--tarih", help="Rapor günü GG.AA.YYYY (varsayılan bugün; örnek veride 14.10.2026)")
    p.add_argument("--ihbar", type=int, default=30, help="İhbar son gününe kaç gün kala uyarılır (varsayılan 30)")
    p.add_argument("--bitis", type=int, default=60, help="Bitişe kaç gün kala uyarılır (varsayılan 60)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "sozlesme_takibi.xlsx")
    a = p.parse_args(argv)
    yol = a.sozlesmeler or ORNEK / "sozlesmeler.csv"
    bugun = tarih(a.tarih) if a.tarih else (date(2026, 10, 14) if a.sozlesmeler is None else date.today())
    if bugun is None or a.ihbar < 0 or a.bitis < 0:
        print("[X] --tarih GG.AA.YYYY; --ihbar ve --bitis 0 ya da üstü olmalı")
        return 2
    if not yol.exists():
        print(f"[X] Dosya bulunamadı: {yol}")
        return 1
    try:
        r = calistir(yol, a.cikti, bugun, a.ihbar, a.bitis)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    o = r["ozet"]
    print(f"[OK] {len(r['sozlesmeler'])} sözleşme · süresi dolan {o['Süresi doldu']} · ihbarı yaklaşan {o['İhbar yaklaşıyor']} · bitişi yaklaşan {o['Bitiş yaklaşıyor']} · ihbar süresi geçen {o['İhbar süresi geçti']}")
    for k in r["kategoriler"]:
        print(f"     {k['kategori']} ({k['pb']}): aylık {k['aylik']:,.2f} · yıllık {k['yillik']:,.2f}")
    for u in r["uyarilar"]:
        if u["onem"] in ("Yüksek", "Orta"):
            print(f"[!] {u['tur']}: {u['kim']} · {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
