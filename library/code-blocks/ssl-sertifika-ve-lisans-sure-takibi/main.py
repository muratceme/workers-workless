"""
SSL Sertifika ve Lisans Süre Takibi — Workers / Workless kod bloğu
Bilgi Teknolojileri › Sistem Uzmanı
Sertifika, alan adı, yazılım lisansı ve destek sözleşmelerinin bitiş tarihlerini izler; yaklaşanları, kullanım fazlası ve atıl lisansları listeler.
  Kalan gün    = bitiş tarihi − rapor günü
  Durum        = Süresi doldu · Kritik (≤ --kritik 7 gün) · Yaklaşıyor (≤ --yakin 30 gün; SSL ve alan adı için) · İzlemede (≤ --izle 90 gün) · Geçerli
  Kullanım %   = kullanılan adet ÷ lisans adedi × 100 (> %100 uyum riski, < --atil %60 atıl lisans)
  Yenileme takvimi = önümüzdeki 12 ayda bitecek varlıkların aylık yenileme maliyeti (para birimi bazında)
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
SUTUN = {"ad": ("varlik", "ad", "alan adi sertifika", "urun", "hizmet"), "tur": ("tur", "kategori", "varlik turu"), "saglayici": ("saglayici", "uretici", "tedarikci", "firma"),
         "bitis": ("bitis tarihi", "bitis", "son gecerlilik", "yenileme tarihi"), "oto": ("otomatik yenileme", "oto yenileme", "auto renew"), "sorumlu": ("sorumlu", "sahip"),
         "maliyet": ("yillik maliyet", "maliyet", "yenileme bedeli", "tutar"), "pb": ("para birimi", "doviz", "pb"), "adet": ("lisans adedi", "adet", "kullanici sayisi"),
         "kullanilan": ("kullanilan", "kullanilan adet", "aktif kullanici")}
SIFIR = Decimal(0)


def _tur(x) -> str:
    k = katla(x)
    if "ssl" in k or "sertifika" in k or "tls" in k:
        return "SSL sertifikası"
    if "alan adi" in k or "domain" in k:
        return "Alan adı"
    if "destek" in k or "bakim" in k or "support" in k:
        return "Destek / bakım"
    if "abonelik" in k or "saas" in k or "bulut" in k:
        return "Abonelik (SaaS)"
    return "Yazılım lisansı" if not k or "lisans" in k else metin(x)


def oku(yol: Path) -> tuple[list[dict], list[dict]]:
    uy, l = [], []
    for r in kayitlar(yol, SUTUN, ("ad", "bitis")):
        ad = metin(r["ad"])
        if not ad:
            continue
        b = tarih(r["bitis"])
        v = {a: (para(r[a]) if metin(r[a]) else None) for a in ("maliyet", "adet", "kullanilan")}
        if b is None or any(metin(r[a]) and (x is None or x < 0) for a, x in v.items()):
            uy.append(uyari("Yüksek", "Satır okunamadı", ad, f"Satır {r['_satir']}: bitiş tarihi, maliyet ya da adetlerden biri okunamadı."))
            continue
        l.append({"ad": ad, "tur": _tur(r["tur"]), "saglayici": metin(r["saglayici"]), "bitis": b, "oto": katla(r["oto"]) in ("evet", "e", "var", "x", "acik"), "sorumlu": metin(r["sorumlu"]),
                  **v, "pb": metin(r["pb"]).upper() or "TRY"})
    return l, uy


def analiz(varliklar: list[dict], bugun: date, kritik: int = 7, yakin: int = 30, izle: int = 90, atil: Decimal = Decimal(60)) -> dict:
    if not varliklar:
        raise ValueError("Okunabilen varlık yok.")
    uy = []
    for v in varliklar:
        v["kalan"] = (v["bitis"] - bugun).days
        kisa = v["tur"] in ("SSL sertifikası", "Alan adı")
        if v["kalan"] < 0:
            v["durum"] = "Süresi doldu"
        elif v["kalan"] <= kritik:
            v["durum"] = "Kritik"
        elif v["kalan"] <= (yakin if kisa else izle):
            v["durum"] = "Yaklaşıyor"
        elif v["kalan"] <= izle:
            v["durum"] = "İzlemede"
        else:
            v["durum"] = "Geçerli"
        v["kullanim"] = yuvarla(v["kullanilan"] * 100 / v["adet"]) if v["adet"] and v["kullanilan"] is not None else None
        kim = f"{v['ad']} ({v['tur']})"
        ek = (" Otomatik yenileme açık; ödeme aracı ve yenilemenin gerçekleştiğini doğrulayın." if v["oto"] else " Otomatik yenileme yok.")
        if v["durum"] == "Süresi doldu":
            uy.append(uyari("Yüksek", "Süresi doldu", kim, f"Bitiş {v['bitis']:%d.%m.%Y} ({-v['kalan']} gün önce)." + (" Sertifika süresi dolan site tarayıcıda güvenlik uyarısı verir." if v["tur"] == "SSL sertifikası" else "")))
        elif v["durum"] == "Kritik":
            uy.append(uyari("Yüksek", "Bitişe çok az kaldı", kim, f"Bitiş {v['bitis']:%d.%m.%Y} ({v['kalan']} gün)." + ek))
        elif v["durum"] == "Yaklaşıyor":
            uy.append(uyari("Orta", "Bitiş yaklaşıyor", kim, f"Bitiş {v['bitis']:%d.%m.%Y} ({v['kalan']} gün)." + ek + (f" Sorumlu: {v['sorumlu']}." if v["sorumlu"] else "")))
        if v["kullanim"] is not None and v["kullanim"] > 100:
            uy.append(uyari("Yüksek", "Lisans adedi aşıldı", kim, f"{v['kullanilan']} kullanım / {v['adet']} lisans (%{v['kullanim']}); lisans uyumsuzluğu riski."))
        elif v["kullanim"] is not None and v["kullanim"] < atil:
            uy.append(uyari("Bilgi", "Atıl lisans", kim, f"{v['kullanilan']} kullanım / {v['adet']} lisans (%{v['kullanim']}); yenilemede adet düşürülebilir."))
        if not v["sorumlu"] and v["durum"] in ("Süresi doldu", "Kritik", "Yaklaşıyor"):
            uy.append(uyari("Bilgi", "Sorumlu yok", kim, "Yenilemeyi takip edecek sorumlu yazılmamış."))
    takvim = {}
    for v in varliklar:
        if 0 <= v["kalan"] <= 365 and v["maliyet"]:
            k = (f"{v['bitis']:%Y-%m}", v["pb"])
            takvim[k] = takvim.get(k, SIFIR) + v["maliyet"]
    sira = {"Süresi doldu": 0, "Kritik": 1, "Yaklaşıyor": 2, "İzlemede": 3, "Geçerli": 4}
    return {"varliklar": sorted(varliklar, key=lambda v: (sira[v["durum"]], v["bitis"], katla(v["ad"]))),
            "ozet": {d: sum(1 for v in varliklar if v["durum"] == d) for d in sira},
            "takvim": [{"ay": a, "pb": p, "tutar": t} for (a, p), t in sorted(takvim.items())], "uyarilar": uy}


def rapor_yaz(cikti: Path, r: dict) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Varlıklar"
    _baslik(ws, ["Varlık", "Tür", "Sağlayıcı", "Bitiş", "Kalan Gün", "Durum", "Otomatik Yenileme", "Sorumlu", "Yıllık Maliyet", "Para Birimi", "Lisans Adedi", "Kullanılan", "Kullanım %"],
            (30, 16, 16, 11, 8, 13, 9, 14, 12, 7, 8, 9, 9), "B2")
    renk = {"Süresi doldu": KIRMIZI, "Kritik": KIRMIZI, "Yaklaşıyor": SARI, "İzlemede": MAVI}
    for v in r["varliklar"]:
        _satir(ws, [v["ad"], v["tur"], v["saglayici"], v["bitis"], v["kalan"], v["durum"], "Evet" if v["oto"] else "Hayır", v["sorumlu"], v["maliyet"], v["pb"], v["adet"], v["kullanilan"],
                    v["kullanim"]], para_s=(9, 13), tarih_s=(4,), dolgu={6: renk.get(v["durum"]), 13: KIRMIZI if v["kullanim"] and v["kullanim"] > 100 else None})
    ws = wb.create_sheet("Yenileme Takvimi")
    _baslik(ws, ["Ay", "Para Birimi", "Yenileme Tutarı"], (10, 10, 16))
    for t in r["takvim"]:
        _satir(ws, [t["ay"], t["pb"], t["tutar"]], para_s=(3,))
    uyari_sayfasi(wb, r["uyarilar"])
    kaydet(wb, cikti)


def calistir(yol: Path, cikti: Path, bugun: date, **k) -> dict:
    v, uy = oku(yol)
    r = analiz(v, bugun, **k)
    r["uyarilar"] = sirala(uy + r["uyarilar"])
    rapor_yaz(cikti, r)
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Sertifika, alan adı ve lisansların bitiş tarihlerini izler, yaklaşanları ve kullanım sorunlarını listeler.")
    p.add_argument("--varliklar", type=Path, help="Varlık, Tür, Sağlayıcı, Bitiş Tarihi, Otomatik Yenileme, Sorumlu, Yıllık Maliyet, Para Birimi, Lisans Adedi, Kullanılan")
    p.add_argument("--tarih", help="Rapor günü GG.AA.YYYY (varsayılan bugün; örnek veride 14.10.2026)")
    p.add_argument("--kritik", type=int, default=7, help="Kritik eşiği, gün (varsayılan 7)")
    p.add_argument("--yakin", type=int, default=30, help="SSL / alan adı için yaklaşma eşiği, gün (varsayılan 30)")
    p.add_argument("--izle", type=int, default=90, help="Lisans ve sözleşmeler için yaklaşma, diğerleri için izleme eşiği, gün (varsayılan 90)")
    p.add_argument("--atil", default="60", help="Kullanım bu yüzdenin altındaysa atıl lisans notu (varsayılan 60)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "sertifika_lisans_takibi.xlsx")
    a = p.parse_args(argv)
    yol = a.varliklar or ORNEK / "varliklar.csv"
    bugun = tarih(a.tarih) if a.tarih else (date(2026, 10, 14) if a.varliklar is None else date.today())
    at = para(a.atil)
    if bugun is None or not 0 <= a.kritik <= a.yakin <= a.izle or at is None or not 0 <= at <= 100:
        print("[X] --tarih GG.AA.YYYY; 0 ≤ --kritik ≤ --yakin ≤ --izle; --atil 0-100 olmalı")
        return 2
    if not yol.exists():
        print(f"[X] Dosya bulunamadı: {yol}")
        return 1
    try:
        r = calistir(yol, a.cikti, bugun, kritik=a.kritik, yakin=a.yakin, izle=a.izle, atil=at)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    o = r["ozet"]
    print(f"[OK] {len(r['varliklar'])} varlık · süresi dolan {o['Süresi doldu']} · kritik {o['Kritik']} · yaklaşan {o['Yaklaşıyor']} · izlemede {o['İzlemede']} · geçerli {o['Geçerli']}")
    for u in r["uyarilar"]:
        if u["onem"] == "Yüksek":
            print(f"[!] {u['tur']}: {u['kim']} · {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
