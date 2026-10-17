"""
Fason Hakediş Hesaplama — Gerisi Bende kod bloğu
Üretim ve Fason Takip › Fason Takip Sorumlusu

Atölyeden dönen sağlam adetler, birim fiyat, eksik ve hatalı ürün kesintilerinden fason hakedişini hesaplar.
  Eksik adet        = verilen − sağlam − defolu                  (eksiye düşerse 0; fazla teslim uyarısı verilir)
  Brüt hakediş      = sağlam adet × fason birim fiyatı           (defolu ve eksik ürüne işçilik ödenmez)
  Tolerans adedi    = verilen × --tolerans %                     (sözleşmedeki fire toleransı; varsayılan 0)
  Kesintiye tabi    = eksik + defolu − tolerans adedi            (eksiyse 0)
  Malzeme kesintisi = kesintiye tabi adet × malzeme bedeli       (kumaş + aksesuar; satırdaki "Malzeme Bedeli")
  KDV               = brüt hakediş × --kdv %;   tevkif edilen KDV = KDV × --tevkifat (ör. 7/10)
  Net ödenecek      = brüt hakediş + KDV − tevkif edilen KDV − malzeme kesintisi − diğer kesinti − avans
KDV ve tevkifat varsayılanı 0'dır; oranları siz verirsiniz. Kesintilerin faturalandırılması ve KDV'si muhasebe uygulamanıza bağlıdır; bu hesap bir ön mutabakattır.
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


SUTUN = {"atolye": ("atolye", "fason", "fasoncu", "tedarikci"), "siparis": ("siparis no", "siparis", "is emri"), "model": ("model", "urun", "aciklama"),
         "verilen": ("verilen adet", "verilen", "gonderilen adet", "kesim adedi"), "saglam": ("saglam adet", "saglam", "teslim edilen saglam", "1 kalite"),
         "defolu": ("defolu adet", "defolu", "hatali adet", "hatali", "2 kalite"), "fiyat": ("birim fiyat", "fason fiyati", "fiyat"),
         "malzeme": ("malzeme bedeli", "birim malzeme bedeli", "kesinti birim bedeli"), "avans": ("avans",), "diger": ("diger kesinti", "ceza", "gecikme cezasi")}
Z = Decimal(0)


def oku(yol: Path) -> tuple[list[dict], list[dict]]:
    uy, l = [], []
    for r in kayitlar(yol, SUTUN, ("atolye", "verilen", "saglam", "fiyat")):
        if not metin(r["atolye"]):
            continue
        kim = f"{metin(r['atolye'])} · {metin(r['siparis'])} {metin(r['model'])}".strip() + f" (satır {r['_satir']})"
        v = {a: para(r[a]) for a in ("verilen", "saglam", "defolu", "fiyat", "malzeme", "avans", "diger")}
        if any(v[a] is None or v[a] < 0 for a in ("verilen", "saglam", "fiyat")):
            uy.append(uyari("Yüksek", "Adet veya fiyat okunamadı", kim, f"Verilen '{metin(r['verilen'])}', sağlam '{metin(r['saglam'])}', fiyat '{metin(r['fiyat'])}'. Satır hakedişe alınmadı."))
            continue
        for a, ad in (("defolu", "Defolu adet"), ("malzeme", "Malzeme bedeli"), ("avans", "Avans"), ("diger", "Diğer kesinti")):
            if v[a] is None and metin(r[a]) or (v[a] is not None and v[a] < 0):
                uy.append(uyari("Yüksek", "Değer okunamadı", kim, f"{ad} '{metin(r[a])}'. 0 sayıldı."))
                v[a] = None
        l.append({"atolye": metin(r["atolye"]), "siparis": metin(r["siparis"]), "model": metin(r["model"]), "kim": kim} | {a: (x if x is not None else (None if a == "malzeme" else Z)) for a, x in v.items()})
    return l, uy


def hesapla(satirlar: list[dict], tolerans: Decimal = Z, kdv: Decimal = Z, tevkifat: Decimal = Z, defolu_esik: Decimal = Decimal(3)) -> dict:
    if not satirlar:
        raise ValueError("Hakedişe alınacak satır yok.")
    uy = []
    for s in satirlar:
        ham = s["verilen"] - s["saglam"] - s["defolu"]
        if ham < 0:
            uy.append(uyari("Yüksek", "Teslim verilenden fazla", s["kim"], f"Sağlam + defolu ({s['saglam'] + s['defolu']:f}) > verilen ({s['verilen']:f}). Verilen adedi kontrol edin; eksik 0 sayıldı."))
        s["eksik"] = max(ham, Z)
        s["brut"] = yuvarla(s["saglam"] * s["fiyat"])
        s["tol_adet"] = s["verilen"] * tolerans / 100
        s["tabi"] = max(s["eksik"] + s["defolu"] - s["tol_adet"], Z)
        if s["tabi"] > 0 and s["malzeme"] is None:
            uy.append(uyari("Orta", "Malzeme bedeli yok", s["kim"], f"{s['tabi']:f} adet kesintiye tabi ama malzeme bedeli girilmemiş; malzeme kesintisi hesaplanmadı."))
        s["kesinti"] = yuvarla(s["tabi"] * (s["malzeme"] or Z))
        s["defolu_oran"] = yuvarla(s["defolu"] / s["verilen"] * 100) if s["verilen"] else None
        s["eksik_oran"] = yuvarla(s["eksik"] / s["verilen"] * 100) if s["verilen"] else None
        s["isaret"] = "Defolu oranı yüksek" if s["defolu_oran"] is not None and s["defolu_oran"] > defolu_esik else ""
    atolye = {}
    for s in satirlar:
        a = atolye.setdefault(katla(s["atolye"]), {"atolye": s["atolye"], "satir": 0} | {k: Z for k in ("verilen", "saglam", "defolu", "eksik", "brut", "kesinti", "diger", "avans")})
        a["satir"] += 1
        for k in ("verilen", "saglam", "defolu", "eksik", "brut", "kesinti", "diger", "avans"):
            a[k] += s[k]
    for a in atolye.values():
        a["kdv"] = yuvarla(a["brut"] * kdv / 100)
        a["tevkif"] = yuvarla(a["kdv"] * tevkifat)
        a["net"] = a["brut"] + a["kdv"] - a["tevkif"] - a["kesinti"] - a["diger"] - a["avans"]
        a["defolu_oran"] = yuvarla(a["defolu"] / a["verilen"] * 100) if a["verilen"] else None
        if a["net"] < 0:
            uy.append(uyari("Yüksek", "Net ödeme eksi", a["atolye"], f"Kesinti ve avanslar hakedişi aşıyor (net {a['net']:f}). Atölye borçlu görünüyor; mutabakat yapın."))
    al = sorted(atolye.values(), key=lambda a: a["atolye"])
    return {"satirlar": satirlar, "atolye": al, "toplam": {k: sum((a[k] for a in al), Z) for k in ("brut", "kdv", "tevkif", "kesinti", "diger", "avans", "net")}, "uyarilar": uy,
            "oran": (tolerans, kdv, tevkifat)}


def rapor_yaz(cikti: Path, r: dict, uy: list[dict]) -> None:
    wb = Workbook()
    tol, kdv, tev = r["oran"]
    ws = wb.active
    ws.title = "Hakediş Özeti"
    _baslik(ws, ["Atölye", "Satır", "Verilen", "Sağlam", "Defolu", "Eksik", "Defolu %", "Brüt Hakediş", f"KDV (%{kdv.normalize():f})", "Tevkif Edilen KDV", "Malzeme Kesintisi", "Diğer Kesinti",
                 "Avans", "Net Ödenecek", "Mutabakat (siz doldurun)"], (26, 7, 10, 10, 9, 9, 9, 16, 14, 14, 15, 13, 13, 16, 24))
    for a in r["atolye"]:
        _satir(ws, [a["atolye"], a["satir"], a["verilen"], a["saglam"], a["defolu"], a["eksik"], a["defolu_oran"], a["brut"], a["kdv"], a["tevkif"], a["kesinti"], a["diger"], a["avans"], a["net"], ""],
               para_s=(7, 8, 9, 10, 11, 12, 13, 14), dolgu={14: KIRMIZI if a["net"] < 0 else None})
    t = r["toplam"]
    _satir(ws, ["Toplam", "", "", "", "", "", "", t["brut"], t["kdv"], t["tevkif"], t["kesinti"], t["diger"], t["avans"], t["net"]], para_s=(8, 9, 10, 11, 12, 13, 14))
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    ws.append([])
    ws.append([f"Fire toleransı %{tol.normalize():f} · KDV %{kdv.normalize():f} · tevkifat oranı {tev.normalize():f}. Net = brüt + KDV − tevkif edilen KDV − malzeme kesintisi − diğer kesinti − avans. Ön mutabakattır."])
    ws = wb.create_sheet("Satırlar")
    _baslik(ws, ["Atölye", "Sipariş", "Model", "Verilen", "Sağlam", "Defolu", "Eksik", "Defolu %", "Eksik %", "Birim Fiyat", "Brüt Hakediş", "Tolerans Adedi", "Kesintiye Tabi Adet",
                 "Malzeme Bedeli", "Malzeme Kesintisi", "Diğer Kesinti", "Avans", "İşaret"], (24, 12, 22, 9, 9, 9, 9, 9, 9, 11, 15, 11, 12, 13, 15, 13, 13, 20))
    for s in r["satirlar"]:
        _satir(ws, [s["atolye"], s["siparis"], s["model"], s["verilen"], s["saglam"], s["defolu"], s["eksik"], s["defolu_oran"], s["eksik_oran"], s["fiyat"], s["brut"], yuvarla(s["tol_adet"]),
                    yuvarla(s["tabi"]), s["malzeme"], s["kesinti"], s["diger"] or None, s["avans"] or None, s["isaret"]], para_s=(8, 9, 10, 11, 12, 13, 14, 15, 16, 17),
               dolgu={18: SARI if s["isaret"] else None, 15: SARI if s["kesinti"] else None})
    uyari_sayfasi(wb, sirala(uy))
    kaydet(wb, cikti)


def calistir(yol: Path, cikti: Path, **oran) -> dict:
    satirlar, uy = oku(yol)
    r = hesapla(satirlar, **oran)
    r["uyarilar"] = sirala(uy + r["uyarilar"])
    rapor_yaz(cikti, r, r["uyarilar"])
    return r


def kesir(s: str) -> Decimal | None:
    s = (s or "").strip()
    try:
        if "/" in s:
            a, b = s.split("/")
            return Decimal(a) / Decimal(b)
        return para(s)
    except (InvalidOperation, ZeroDivisionError, ValueError):
        return None


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Fason atölye hakedişini sağlam adet, birim fiyat ve kesintilerden hesaplar.")
    p.add_argument("--teslimler", type=Path, default=ORNEK / "fason_teslimler.csv",
                   help="Atölye, Sipariş No, Model, Verilen Adet, Sağlam Adet, Defolu Adet, Birim Fiyat, Malzeme Bedeli, Avans, Diğer Kesinti")
    p.add_argument("--tolerans", help="Sözleşmedeki fire toleransı %% (varsayılan 0; örnek veride 1)")
    p.add_argument("--kdv", help="KDV oranı %% (varsayılan 0; oranı siz verin)")
    p.add_argument("--tevkifat", help="KDV tevkifat oranı, ör. 7/10 (varsayılan 0)")
    p.add_argument("--defolu-esik", default="3", help="Bu defolu %%'sinin üstü işaretlenir (varsayılan 3)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "fason_hakedis.xlsx")
    a = p.parse_args(argv)
    if not a.teslimler.exists():
        print(f"[X] Dosya bulunamadı: {a.teslimler}")
        return 1
    ornek = a.teslimler == ORNEK / "fason_teslimler.csv"
    tol = para(a.tolerans) if a.tolerans is not None else Decimal(1) if ornek else Z
    kdv = para(a.kdv) if a.kdv is not None else Decimal(10) if ornek else Z
    tev = kesir(a.tevkifat) if a.tevkifat is not None else Decimal(7) / 10 if ornek else Z
    de = para(a.defolu_esik)
    if tol is None or not 0 <= tol <= 100 or kdv is None or not 0 <= kdv <= 100 or tev is None or not 0 <= tev <= 1 or de is None or de < 0:
        print("[X] --tolerans ve --kdv 0-100 arasında; --tevkifat 0 ile 1 arasında bir kesir (ör. 7/10); --defolu-esik sıfır veya pozitif olmalı")
        return 2
    try:
        r = calistir(a.teslimler, a.cikti, tolerans=tol, kdv=kdv, tevkifat=tev, defolu_esik=de)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    for x in r["atolye"]:
        print(f"[OK] {x['atolye']}: brüt {x['brut']:,.2f} · malzeme kesintisi {x['kesinti']:,.2f} · avans {x['avans']:,.2f} · net ödenecek {x['net']:,.2f}")
    for u in r["uyarilar"]:
        if u["onem"] == "Yüksek":
            print(f"[!] {u['kim']} · {u['tur']}: {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
