"""
Mağaza Vardiya Planlama — Workers / Workless kod bloğu
Mağaza Operasyonları › Mağaza Müdür Yardımcısı

Saatlik müşteri trafiği tahmini, personel izinleri ve çalışma süresi kurallarına göre haftalık vardiya planı çıkarır.
  Saatlik ihtiyaç  = en büyük(--min-personel, beklenen müşteri ÷ --musteri-basina, yukarı yuvarlanır)
  Vardiya net süre = (bitiş − başlangıç) − mola                        (ara dinlenmesi çalışma süresinden sayılmaz)
  Atama (sezgisel) = günler en yoğundan başlanarak ele alınır; her adımda karşılanmamış saatleri en çok kapatan vardiya seçilir ve
                     kurallara uyan, haftalık kalan saati en çok olan personele verilir (eşitlikte izin günü çok olana); ihtiyaç bitene veya uygun personel kalmayana kadar sürer
  Kurallar         = izin gününde atama yok · günde tek vardiya · haftalık net saat ≤ personelin sınırı (--haftalik-saat, 45) · vardiya net süresi ≤ --gunluk-saat (11)
                     iki vardiya arası dinlenme ≥ --dinlenme-saat (11) · haftada en çok --max-gun (6) gün çalışma
  Mola kontrolü    = vardiyadaki mola; süre 4 saate kadar 15 dk, 4 saatten 7,5 saate kadar 30 dk, 7,5 saatten fazla 60 dk'nın altındaysa uyarı verilir
Varsayılan sınırlar İş Kanunu'ndaki genel çalışma süresi kurallarına göre seçilmiştir; güncel mevzuatı, toplu iş sözleşmesini ve iş sözleşmelerini siz doğrulayın.
Plan sezgiseldir (en iyi çözümü garanti etmez) ve taslaktır; fazla çalışma, resmi tatil, gece çalışması, genç ve gebe çalışan kuralları hesaba katılmaz.
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


TRAFIK = {"gun": ("gun", "tarih"), "saat": ("saat", "saat dilimi", "baslangic saati"), "musteri": ("musteri", "beklenen musteri", "musteri sayisi", "ziyaretci", "trafik")}
PERSONEL = {"ad": ("personel", "ad", "ad soyad", "calisan"), "limit": ("haftalik saat", "haftalik calisma saati", "sozlesme saati"), "izin": ("izin gunleri", "izin", "izinli gunler", "hafta tatili")}
VARDIYA = {"ad": ("vardiya", "vardiya adi"), "basla": ("baslangic", "baslama", "giris"), "bitis": ("bitis", "cikis"), "mola": ("mola", "mola dk", "ara dinlenmesi")}
GUNLER = ("Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar")
Z = Decimal(0)


def gun_coz(x) -> int | None:
    k, adlar = katla(x), [katla(g) for g in GUNLER]
    if k in adlar:
        return adlar.index(k)
    kisaltma = [i for i, g in enumerate(adlar) if len(k) >= 3 and g.startswith(k)]      # "Sal", "Çar", "Per" tanınır; "Paz" ve "Cum" iki güne uyduğu için tanınmaz
    if len(kisaltma) == 1:
        return kisaltma[0]
    t = tarih(x)
    return t.weekday() if t else None


def saat_coz(x) -> Decimal | None:
    """10 · 10:30 · 10.30 · saat nesnesi → ondalık saat (10,5)"""
    if hasattr(x, "hour"):
        return Decimal(x.hour) + Decimal(x.minute) / 60
    m = re.fullmatch(r"(\d{1,2})(?:[:.](\d{2}))?(?::\d{2})?", metin(x))
    if not m or int(m[1]) > 24 or int(m[2] or 0) > 59:
        return None
    return Decimal(int(m[1])) + Decimal(int(m[2] or 0)) / 60


def saat_yaz(s: Decimal) -> str:
    return f"{int(s):02d}:{int((s - int(s)) * 60):02d}"


def asgari_mola(brut: Decimal) -> int:
    return 15 if brut <= 4 else 30 if brut <= Decimal("7.5") else 60


def oku(trafik_yol: Path, personel_yol: Path, vardiya_yol: Path, haftalik: Decimal) -> tuple[dict, list[dict], list[dict], list[dict]]:
    uy, trafik, personel, vardiyalar = [], defaultdict(lambda: defaultdict(lambda: Z)), [], []
    for r in kayitlar(trafik_yol, TRAFIK, ("gun", "saat", "musteri")):
        if not metin(r["gun"]):
            continue
        g, s, m = gun_coz(r["gun"]), saat_coz(r["saat"]), para(r["musteri"])
        if g is None or s is None or m is None or m < 0:
            uy.append(uyari("Yüksek", "Trafik satırı okunamadı", f"{trafik_yol.name} satır {r['_satir']}", f"Gün '{metin(r['gun'])}', saat '{metin(r['saat'])}', müşteri '{metin(r['musteri'])}'. Satır alınmadı."))
            continue
        trafik[g][int(s)] += m
    for r in kayitlar(personel_yol, PERSONEL, ("ad",)):
        ad = metin(r["ad"])
        if not ad:
            continue
        limit, izin = para(r["limit"]) if metin(r["limit"]) else haftalik, set()
        if limit is None or limit <= 0:
            uy.append(uyari("Yüksek", "Haftalık saat okunamadı", ad, f"'{metin(r['limit'])}'. Varsayılan {haftalik:f} saat kullanıldı."))
            limit = haftalik
        elif limit > haftalik:
            uy.append(uyari("Orta", "Haftalık saat sınırın üzerinde", ad, f"{limit:f} saat > {haftalik:f} saat. Sınır değeri kullanıldı."))
            limit = haftalik
        for parca in re.split(r"[,;/]", metin(r["izin"])):
            if parca.strip():
                g = gun_coz(parca)
                if g is None:
                    uy.append(uyari("Yüksek", "İzin günü tanınmadı", ad, f"'{parca.strip()}'. Gün adı (Pazartesi … Pazar) veya tarih yazın; bu izin plana yansımadı."))
                else:
                    izin.add(g)
        personel.append({"ad": ad, "limit": limit, "izin": izin})
    for r in kayitlar(vardiya_yol, VARDIYA, ("ad", "basla", "bitis")):
        ad = metin(r["ad"])
        if not ad:
            continue
        b, s, mola = saat_coz(r["basla"]), saat_coz(r["bitis"]), para(r["mola"]) if metin(r["mola"]) else Z
        if b is None or s is None or mola is None or mola < 0:
            uy.append(uyari("Yüksek", "Vardiya okunamadı", ad, f"Başlangıç '{metin(r['basla'])}', bitiş '{metin(r['bitis'])}', mola '{metin(r['mola'])}'. Vardiya kullanılmadı."))
            continue
        if s <= b:
            uy.append(uyari("Yüksek", "Gece yarısını aşan vardiya", ad, f"{saat_yaz(b)}-{saat_yaz(s)}: bitiş başlangıçtan sonra olmalı. Gece yarısını aşan vardiyalar desteklenmez; vardiya kullanılmadı."))
            continue
        net = s - b - mola / 60
        if net <= 0:
            uy.append(uyari("Yüksek", "Vardiya okunamadı", ad, "Mola vardiya süresine eşit veya daha uzun. Vardiya kullanılmadı."))
            continue
        if mola < asgari_mola(s - b):
            uy.append(uyari("Orta", "Mola asgari sürenin altında", ad, f"{saat_yaz(b)}-{saat_yaz(s)} ({s - b:f} saat) için mola {mola:f} dk; bu süre için en az {asgari_mola(s - b)} dk ara dinlenmesi beklenir. Mevzuatı kontrol edin."))
        vardiyalar.append({"ad": ad, "basla": b, "bitis": s, "mola": mola, "net": net, "etiket": f"{ad} {saat_yaz(b)}-{saat_yaz(s)}"})
    return {g: dict(v) for g, v in trafik.items()}, personel, vardiyalar, uy


def planla(trafik: dict, personel: list[dict], vardiyalar: list[dict], musteri_basina: Decimal = Decimal(20), min_personel: int = 2, gunluk: Decimal = Decimal(11), dinlenme: Decimal = Decimal(11),
           max_gun: int = 6) -> dict:
    if not trafik:
        raise ValueError("Trafik verisi yok.")
    if not personel:
        raise ValueError("Personel listesi boş.")
    kullanilan = [v for v in vardiyalar if v["net"] <= gunluk]
    if not kullanilan:
        raise ValueError("Kullanılabilir vardiya yok (hepsi günlük çalışma sınırının üzerinde veya okunamadı).")
    ihtiyac = {g: {h: max(min_personel, math.ceil(m / musteri_basina)) for h, m in saatler.items()} for g, saatler in trafik.items()}
    for p in personel:
        p["saat"], p["gun"] = Z, 0
    atama, atanan = {}, {g: {h: 0 for h in ihtiyac[g]} for g in ihtiyac}

    def kapsar(v: dict, h: int) -> bool:
        return v["basla"] <= h and v["bitis"] >= h + 1

    def uygun(i: int, g: int, v: dict) -> bool:
        p = personel[i]
        if g in p["izin"] or (i, g) in atama or p["saat"] + v["net"] > p["limit"] or p["gun"] >= max_gun:
            return False
        once, sonra = atama.get((i, g - 1)), atama.get((i, g + 1))
        return not (once and 24 - once["bitis"] + v["basla"] < dinlenme) and not (sonra and 24 - v["bitis"] + sonra["basla"] < dinlenme)

    for g in sorted(ihtiyac, key=lambda g: (-sum(ihtiyac[g].values()), g)):
        while True:
            kalan = {h for h in ihtiyac[g] if atanan[g][h] < ihtiyac[g][h]}
            secenek = sorted(((sum(kapsar(v, h) for h in kalan), j) for j, v in enumerate(kullanilan)), key=lambda x: (-x[0], kullanilan[x[1]]["net"], x[1]))
            for kazanc, j in secenek:
                v = kullanilan[j]
                adaylar = [i for i in range(len(personel)) if uygun(i, g, v)] if kazanc else []
                if adaylar:
                    i = max(adaylar, key=lambda i: (personel[i]["limit"] - personel[i]["saat"], -personel[i]["gun"], len(personel[i]["izin"]), -i))
                    atama[i, g] = v
                    personel[i]["saat"] += v["net"]
                    personel[i]["gun"] += 1
                    for h in atanan[g]:
                        atanan[g][h] += kapsar(v, h)
                    break
            else:
                break
    gunler = sorted(ihtiyac)
    kapsama = [{"gun": GUNLER[g], "g": g, "saat": h, "musteri": trafik[g][h], "ihtiyac": ihtiyac[g][h], "atanan": atanan[g][h], "fark": atanan[g][h] - ihtiyac[g][h]} for g in gunler for h in sorted(ihtiyac[g])]
    ozet = []
    for g in gunler:
        k = [x for x in kapsama if x["g"] == g]
        ozet.append({"gun": GUNLER[g], "ihtiyac": sum(x["ihtiyac"] for x in k), "atanan": sum(x["atanan"] for x in k), "eksik": sum(-x["fark"] for x in k if x["fark"] < 0),
                     "fazla": sum(x["fark"] for x in k if x["fark"] > 0), "eksik_saat": [x["saat"] for x in k if x["fark"] < 0], "vardiya": sum(gg == g for _, gg in atama)})
    plan = [{"personel": p["ad"], "limit": p["limit"], "saat": p["saat"], "gun": p["gun"], "kalan": p["limit"] - p["saat"],
             "gunler": {g: atama[i, g]["etiket"] if (i, g) in atama else "İzin" if g in p["izin"] else "" for g in gunler}} for i, p in enumerate(personel)]
    return {"plan": plan, "kapsama": kapsama, "ozet": ozet, "gunler": gunler, "atama": atama, "atlanan_vardiya": [v["ad"] for v in vardiyalar if v["net"] > gunluk],
            "toplam": {"vardiya": len(atama), "ihtiyac": sum(o["ihtiyac"] for o in ozet), "atanan": sum(o["atanan"] for o in ozet), "eksik": sum(o["eksik"] for o in ozet), "fazla": sum(o["fazla"] for o in ozet),
                       "saat": sum((p["saat"] for p in personel), Z)}, "ayar": (musteri_basina, min_personel, gunluk, dinlenme, max_gun)}


def rapor_yaz(cikti: Path, r: dict, uy: list[dict]) -> None:
    wb = Workbook()
    mb, mp, gs, ds, mg = r["ayar"]
    ws = wb.active
    ws.title = "Vardiya Planı"
    _baslik(ws, ["Personel", *[GUNLER[g] for g in r["gunler"]], "Toplam Net Saat", "Haftalık Sınır", "Kalan Saat", "Çalışılan Gün"], (24, *[20] * len(r["gunler"]), 10, 10, 9, 9), dondur="B2")
    for p in r["plan"]:
        hucre = [p["gunler"][g] for g in r["gunler"]]
        _satir(ws, [p["personel"], *hucre, p["saat"], p["limit"], p["kalan"], p["gun"]], para_s=tuple(range(len(hucre) + 2, len(hucre) + 5)),
               dolgu={j: (SARI if h == "İzin" else MAVI if h else None) for j, h in enumerate(hucre, 2)})
    ws.append([])
    ws.append([f"Saatlik ihtiyaç = en büyük({mp}, müşteri ÷ {mb.normalize():f}). Kurallar: günde tek vardiya, vardiya net ≤ {gs.normalize():f} saat, vardiyalar arası dinlenme ≥ {ds.normalize():f} saat, haftada en çok {mg} gün. "
               "Plan sezgisel bir taslaktır; mevzuata ve sözleşmelere uygunluğunu kontrol edip onaylayın."])
    ws = wb.create_sheet("Gün Özeti")
    _baslik(ws, ["Gün", "Atanan Vardiya", "İhtiyaç (adam-saat)", "Atanan (adam-saat)", "Eksik (adam-saat)", "Fazla (adam-saat)", "Eksik Kalan Saatler"], (14, 10, 12, 12, 12, 12, 50))
    for o in r["ozet"]:
        _satir(ws, [o["gun"], o["vardiya"], o["ihtiyac"], o["atanan"], o["eksik"] or None, o["fazla"] or None, ", ".join(f"{h:02d}:00" for h in o["eksik_saat"])], dolgu={5: KIRMIZI if o["eksik"] else None}, sar=(7,))
    t = r["toplam"]
    _satir(ws, ["Toplam", t["vardiya"], t["ihtiyac"], t["atanan"], t["eksik"] or None, t["fazla"] or None, ""])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    ws = wb.create_sheet("Saatlik Kapsama")
    _baslik(ws, ["Gün", "Saat", "Beklenen Müşteri", "İhtiyaç", "Atanan", "Fark"], (14, 12, 12, 9, 9, 9))
    for k in r["kapsama"]:
        _satir(ws, [k["gun"], f"{k['saat']:02d}:00-{k['saat'] + 1:02d}:00", k["musteri"], k["ihtiyac"], k["atanan"], k["fark"]], dolgu={6: KIRMIZI if k["fark"] < 0 else MAVI if k["fark"] > 0 else None})
    uyari_sayfasi(wb, sirala(uy))
    kaydet(wb, cikti)


def calistir(trafik_yol: Path, personel_yol: Path, vardiya_yol: Path, cikti: Path, haftalik: Decimal = Decimal(45), **ayar) -> dict:
    trafik, personel, vardiyalar, uy = oku(trafik_yol, personel_yol, vardiya_yol, haftalik)
    r = planla(trafik, personel, vardiyalar, **ayar)
    for v in r["atlanan_vardiya"]:
        uy.append(uyari("Yüksek", "Vardiya günlük sınırın üzerinde", v, f"Net süresi {r['ayar'][2]:f} saati aşıyor; planlamada kullanılmadı."))
    for o in r["ozet"]:
        if o["eksik"]:
            uy.append(uyari("Yüksek", "Karşılanamayan ihtiyaç", o["gun"], f"{o['eksik']} adam-saat eksik ({', '.join(f'{h:02d}:00' for h in o['eksik_saat'])}). Uygun personel kalmadı: izinleri, haftalık saatleri veya vardiya "
                            "kalıplarını gözden geçirin."))
    r["uyarilar"] = sirala(uy)
    rapor_yaz(cikti, r, uy)
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Trafik tahmini, izinler ve çalışma süresi kurallarına göre haftalık vardiya planı taslağı çıkarır.")
    p.add_argument("--trafik", type=Path, default=ORNEK / "saatlik_trafik.csv", help="Gün, Saat, Müşteri (beklenen)")
    p.add_argument("--personel", type=Path, default=ORNEK / "personel.csv", help="Personel, Haftalık Saat, İzin Günleri (virgülle)")
    p.add_argument("--vardiyalar", type=Path, default=ORNEK / "vardiyalar.csv", help="Vardiya, Başlangıç, Bitiş, Mola (dk)")
    p.add_argument("--musteri-basina", default="20", help="Bir personelin bir saatte karşılayabildiği müşteri sayısı (varsayılan 20; mağazanıza göre belirleyin)")
    p.add_argument("--min-personel", default="2", help="Her açık saatte bulunması gereken en az personel (varsayılan 2)")
    p.add_argument("--haftalik-saat", default="45", help="Haftalık net çalışma saati üst sınırı (varsayılan 45)")
    p.add_argument("--gunluk-saat", default="11", help="Günlük net çalışma saati üst sınırı (varsayılan 11)")
    p.add_argument("--dinlenme-saat", default="11", help="İki vardiya arasındaki en az dinlenme (varsayılan 11 saat)")
    p.add_argument("--max-gun", default="6", help="Haftada en çok çalışılacak gün (varsayılan 6)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "vardiya_plani.xlsx")
    a = p.parse_args(argv)
    d = [para(x) for x in (a.musteri_basina, a.min_personel, a.haftalik_saat, a.gunluk_saat, a.dinlenme_saat, a.max_gun)]
    if any(x is None or x < 0 for x in d) or d[0] == 0 or d[2] == 0 or d[3] == 0 or d[1] != int(d[1]) or d[5] != int(d[5]) or not 1 <= d[5] <= 7:
        print("[X] --musteri-basina, --haftalik-saat ve --gunluk-saat pozitif; --min-personel tam sayı; --max-gun 1-7 arasında tam sayı olmalı")
        return 2
    for y in (a.trafik, a.personel, a.vardiyalar):
        if not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    try:
        r = calistir(a.trafik, a.personel, a.vardiyalar, a.cikti, d[2], musteri_basina=d[0], min_personel=int(d[1]), gunluk=d[3], dinlenme=d[4], max_gun=int(d[5]))
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    t = r["toplam"]
    print(f"[OK] {len(r['plan'])} personel · {t['vardiya']} vardiya · {t['saat']:f} net saat · ihtiyaç {t['ihtiyac']} adam-saat, atanan {t['atanan']}")
    print(f"[{'!' if t['eksik'] else 'OK'}] Eksik {t['eksik']} adam-saat · fazla {t['fazla']} adam-saat")
    for o in r["ozet"]:
        if o["eksik"]:
            print(f"[!] {o['gun']}: {o['eksik']} adam-saat eksik ({', '.join(f'{h:02d}:00' for h in o['eksik_saat'])})")
    print(f"[OK] Plan taslağı: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
