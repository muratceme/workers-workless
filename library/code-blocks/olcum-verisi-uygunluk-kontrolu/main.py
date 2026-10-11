"""
Ölçüm Verisi Uygunluk Kontrolü — Workers / Workless kod bloğu
Kalite Güvence ve Kalite Kontrol › Kalite Kontrol Elemanı

Ölçüm sonuçlarını teknik resim toleranslarıyla karşılaştırır; uygunsuz ve sınıra yakın ölçüleri işaretler.
  Alt limit = nominal + alt tolerans (işaretiyle; ör. −0,021)     Üst limit = nominal + üst tolerans
  (Alt Limit / Üst Limit sütunları doluysa doğrudan onlar kullanılır; tek taraflı tolerans için yalnız biri girilir)
  Uygun        = alt limit ≤ ölçülen ≤ üst limit        (limit değerinin kendisi uygundur)
  Sınıra yakın = uygun, ama en yakın limite uzaklığı tolerans aralığının --sinir (varsayılan %10) kadarı veya daha azı (çift taraflı toleransta)
  Parça sonucu = ölçülerinden biri uygunsuzsa RET
Sayılarda virgül de nokta da ondalık ayırıcıdır (25,000 = 25.000 = yirmi beş). Ölçüm belirsizliği hesaba katılmaz.
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


TOL_SUTUN = {"parca": ("parca kodu", "parca", "urun kodu", "resim no"), "kar": ("karakteristik", "olcu", "olcu adi", "balon no", "ozellik"), "nominal": ("nominal", "nominal deger"),
             "alt_tol": ("alt tolerans", "alt sapma"), "ust_tol": ("ust tolerans", "ust sapma"), "alt": ("alt limit", "min", "lsl"), "ust": ("ust limit", "max", "maks", "usl"),
             "birim": ("birim",)}
OLC_SUTUN = {"parca": ("parca kodu", "parca", "urun kodu", "resim no"), "numune": ("numune no", "seri no", "numune", "parca no", "sira"), "kar": ("karakteristik", "olcu", "olcu adi", "balon no", "ozellik"),
             "deger": ("olculen", "olculen deger", "olcum", "deger", "sonuc"), "tarih": ("tarih",), "operator": ("operator", "olcen"), "cihaz": ("cihaz", "olcum cihazi")}


def olcu(x) -> Decimal | None:
    """Ölçü değeri: virgül ve nokta ondalık ayırıcıdır; binlik ayırıcı kabul edilmez."""
    if x in (None, ""):
        return None
    if isinstance(x, (int, float, Decimal)):
        return Decimal(str(x))
    s = re.sub(r"[^\d,.\-+]", "", str(x).replace("−", "-").replace("–", "-")).replace(",", ".")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def tolerans_oku(yol: Path) -> tuple[dict, list[dict]]:
    tol, uy = {}, []
    for r in kayitlar(yol, TOL_SUTUN, ("kar",)):
        kar = metin(r["kar"])
        if not kar:
            continue
        kim = f"{metin(r['parca'])} · {kar}".strip(" ·")
        n, at, ut, alt, ust = (olcu(r[a]) for a in ("nominal", "alt_tol", "ust_tol", "alt", "ust"))
        if alt is None and n is not None and at is not None:
            alt = n + at
        if ust is None and n is not None and ut is not None:
            ust = n + ut
        if alt is None and ust is None:
            uy.append(uyari("Yüksek", "Tolerans tanımı eksik", kim, f"{yol.name} satır {r['_satir']}: nominal + alt / üst tolerans ya da alt / üst limit girilmeli. Karakteristik kontrol edilemedi."))
            continue
        if alt is not None and ust is not None and alt > ust:
            uy.append(uyari("Yüksek", "Alt limit üst limitten büyük", kim, f"{yol.name} satır {r['_satir']}: alt {alt}, üst {ust}. Alt toleransı eksi işaretiyle yazın (ör. -0,05). Karakteristik kontrol edilemedi."))
            continue
        if alt is not None and alt == ust:
            uy.append(uyari("Orta", "Tolerans aralığı sıfır", kim, f"{yol.name} satır {r['_satir']}: alt ve üst limit aynı ({alt}). ± tolerans demek istediyseniz alt toleransı eksi işaretiyle yazın."))
        a = (katla(r["parca"]), katla(kar))
        if a in tol:
            uy.append(uyari("Yüksek", "Karakteristik tekrar ediyor", kim, f"{yol.name} satır {r['_satir']}: aynı parça ve karakteristik iki kez tanımlı; ilk satır kullanıldı."))
            continue
        tol[a] = {"parca": metin(r["parca"]), "kar": kar, "nominal": n, "alt": alt, "ust": ust, "birim": metin(r["birim"])}
    if not tol:
        raise ValueError(f"{yol.name}: kullanılabilir tolerans satırı yok.")
    return tol, uy


def olcum_oku(yol: Path) -> tuple[list[dict], list[dict]]:
    l, uy = [], []
    for r in kayitlar(yol, OLC_SUTUN, ("kar", "deger")):
        if not metin(r["kar"]):
            continue
        v = olcu(r["deger"])
        if v is None:
            uy.append(uyari("Yüksek", "Ölçüm değeri okunamadı", f"{metin(r['numune'])} · {metin(r['kar'])}", f"{yol.name} satır {r['_satir']}: '{metin(r['deger'])}'. Ölçüm alınmadı."))
            continue
        l.append({"parca": metin(r["parca"]), "numune": metin(r["numune"]) or "—", "kar": metin(r["kar"]), "deger": v, "tarih": tarih(r["tarih"]), "operator": metin(r["operator"]),
                  "cihaz": metin(r["cihaz"])})
    return l, uy


def degerlendir(v: Decimal, t: dict, sinir: Decimal = Decimal(10)) -> tuple[str, Decimal | None]:
    """(sonuç, limit dışına taşma miktarı)"""
    if t["ust"] is not None and v > t["ust"]:
        return "Uygunsuz (üst)", v - t["ust"]
    if t["alt"] is not None and v < t["alt"]:
        return "Uygunsuz (alt)", t["alt"] - v
    if t["alt"] is not None and t["ust"] is not None and min(v - t["alt"], t["ust"] - v) <= (t["ust"] - t["alt"]) * sinir / 100:
        return "Sınıra yakın", None
    return "Uygun", None


def kontrol(olcumler: list[dict], tol: dict, sinir: Decimal = Decimal(10)) -> dict:
    uy, sonuc, eslesmeyen = [], [], defaultdict(int)
    for o in olcumler:
        t = tol.get((katla(o["parca"]), katla(o["kar"]))) or tol.get(("", katla(o["kar"])))
        if not t:
            eslesmeyen[(o["parca"], o["kar"])] += 1
            continue
        s, tasma = degerlendir(o["deger"], t, sinir)
        sonuc.append(o | {"t": t, "sonuc": s, "tasma": tasma, "sapma": o["deger"] - t["nominal"] if t["nominal"] is not None else None})
    for (p, k), n in eslesmeyen.items():
        uy.append(uyari("Yüksek", "Ölçümün toleransı tanımsız", f"{p} · {k}".strip(" ·"), f"{n} ölçüm için tolerans dosyasında karşılık yok; kontrol edilmedi."))
    kar = {}
    for s in sonuc:
        x = kar.setdefault((katla(s["parca"]), katla(s["kar"])), {"parca": s["parca"], "kar": s["kar"], "t": s["t"], "d": [], "uygunsuz": 0, "yakin": 0})
        x["d"].append(s["deger"])
        x["uygunsuz"] += s["sonuc"].startswith("Uygunsuz")
        x["yakin"] += s["sonuc"] == "Sınıra yakın"
    for x in kar.values():
        x["n"], x["min"], x["max"] = len(x["d"]), min(x["d"]), max(x["d"])
        x["ort"] = yuvarla(sum(x["d"]) / len(x["d"]), "0.0001")
        x["oran"] = yuvarla(Decimal(x["uygunsuz"]) / x["n"] * 100)
    olculen = {(katla(s["parca"]), katla(s["kar"])) for s in sonuc}
    for a, t in tol.items():
        if a not in olculen and not (a[0] == "" and any(k == a[1] for _, k in olculen)):
            uy.append(uyari("Orta", "Karakteristik ölçülmemiş", f"{t['parca']} · {t['kar']}".strip(" ·"), "Tolerans dosyasında var, ölçüm dosyasında hiç ölçümü yok."))
    parca = {}
    for s in sonuc:
        x = parca.setdefault((s["parca"], s["numune"]), {"parca": s["parca"], "numune": s["numune"], "olcu": 0, "uygunsuz": [], "yakin": []})
        x["olcu"] += 1
        if s["sonuc"].startswith("Uygunsuz"):
            x["uygunsuz"].append(s["kar"])
        elif s["sonuc"] == "Sınıra yakın":
            x["yakin"].append(s["kar"])
    for x in parca.values():
        x["sonuc"] = "RET" if x["uygunsuz"] else "KABUL"
    return {"sonuc": sonuc, "kar": sorted(kar.values(), key=lambda x: (-x["uygunsuz"], x["parca"], x["kar"])), "parca": list(parca.values()), "uyarilar": uy,
            "say": {k: sum(s["sonuc"].startswith(k) for s in sonuc) for k in ("Uygun", "Sınıra yakın", "Uygunsuz")} | {"Uygun": sum(s["sonuc"] == "Uygun" for s in sonuc)}}


def rapor_yaz(cikti: Path, r: dict, uy: list[dict]) -> None:
    wb = Workbook()
    renk = {"Uygun": None, "Sınıra yakın": SARI, "Uygunsuz (üst)": KIRMIZI, "Uygunsuz (alt)": KIRMIZI}
    sf = "0.000#"

    def bicim(ws, sutunlar):
        for j in sutunlar:
            ws.cell(ws.max_row, j).number_format = sf

    ws = wb.active
    ws.title = "Parça Özeti"
    _baslik(ws, ["Parça Kodu", "Numune No", "Ölçü Sayısı", "Sonuç", "Uygunsuz Ölçüler", "Sınıra Yakın Ölçüler", "Karar / Not (siz doldurun)"], (18, 14, 9, 10, 44, 44, 30))
    for x in sorted(r["parca"], key=lambda x: x["sonuc"] != "RET"):
        _satir(ws, [x["parca"], x["numune"], x["olcu"], x["sonuc"], ", ".join(x["uygunsuz"]), ", ".join(x["yakin"]), ""], dolgu={4: KIRMIZI if x["sonuc"] == "RET" else YESIL}, sar=(5, 6))
    ws.append([])
    ws.append(["Ölçüm belirsizliği hesaba katılmamıştır. Limit değerinin kendisi uygun sayılır."])
    ws = wb.create_sheet("Karakteristik Özeti")
    _baslik(ws, ["Parça Kodu", "Karakteristik", "Birim", "Nominal", "Alt Limit", "Üst Limit", "Ölçüm Sayısı", "En Küçük", "En Büyük", "Ortalama", "Uygunsuz", "Uygunsuz %", "Sınıra Yakın"],
            (18, 26, 8, 12, 12, 12, 9, 12, 12, 12, 9, 10, 9))
    for x in r["kar"]:
        t = x["t"]
        _satir(ws, [x["parca"], x["kar"], t["birim"], t["nominal"], t["alt"], t["ust"], x["n"], x["min"], x["max"], x["ort"], x["uygunsuz"] or None, x["oran"], x["yakin"] or None],
               para_s=(12,), dolgu={11: KIRMIZI if x["uygunsuz"] else None, 13: SARI if x["yakin"] else None})
        bicim(ws, (4, 5, 6, 8, 9, 10))
    ws = wb.create_sheet("Ölçüm Sonuçları")
    _baslik(ws, ["Parça Kodu", "Numune No", "Karakteristik", "Ölçülen", "Nominal", "Alt Limit", "Üst Limit", "Nominalden Sapma", "Sonuç", "Limit Dışı Miktar", "Tarih", "Operatör", "Cihaz"],
            (18, 14, 26, 12, 12, 12, 12, 12, 16, 12, 12, 16, 16))
    for s in sorted(r["sonuc"], key=lambda s: (not s["sonuc"].startswith("Uygunsuz"), s["sonuc"] != "Sınıra yakın")):
        t = s["t"]
        _satir(ws, [s["parca"], s["numune"], s["kar"], s["deger"], t["nominal"], t["alt"], t["ust"], s["sapma"], s["sonuc"], s["tasma"], s["tarih"], s["operator"], s["cihaz"]], tarih_s=(11,),
               dolgu={9: renk[s["sonuc"]]})
        bicim(ws, (4, 5, 6, 7, 8, 10))
    uyari_sayfasi(wb, sirala(uy))
    kaydet(wb, cikti)


def calistir(olcum_yol: Path, tol_yol: Path, cikti: Path, sinir: Decimal = Decimal(10)) -> dict:
    tol, uy1 = tolerans_oku(tol_yol)
    olcumler, uy2 = olcum_oku(olcum_yol)
    if not olcumler:
        raise ValueError(f"{olcum_yol.name}: okunabilen ölçüm yok.")
    r = kontrol(olcumler, tol, sinir)
    r["uyarilar"] = sirala(uy1 + uy2 + r["uyarilar"])
    rapor_yaz(cikti, r, r["uyarilar"])
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Ölçüm sonuçlarını teknik resim toleranslarıyla karşılaştırır.")
    p.add_argument("--olcumler", type=Path, default=ORNEK / "olcumler.csv", help="Parça Kodu, Numune No, Karakteristik, Ölçülen, Tarih, Operatör, Cihaz")
    p.add_argument("--toleranslar", type=Path, default=ORNEK / "toleranslar.csv", help="Parça Kodu, Karakteristik, Nominal, Alt Tolerans, Üst Tolerans veya Alt Limit, Üst Limit, Birim")
    p.add_argument("--sinir", default="10", help="Sınıra yakın sayılacak pay: tolerans aralığının %%'si (varsayılan 10; 0 = kapalı)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "olcum_uygunluk.xlsx")
    a = p.parse_args(argv)
    for y in (a.olcumler, a.toleranslar):
        if not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    sinir = para(a.sinir)
    if sinir is None or not 0 <= sinir < 50:
        print("[X] --sinir 0 ile 50 arasında bir yüzde olmalı")
        return 2
    try:
        r = calistir(a.olcumler, a.toleranslar, a.cikti, sinir)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    s = r["say"]
    print(f"[OK] {len(r['sonuc'])} ölçüm · {s['Uygun']} uygun · {s['Sınıra yakın']} sınıra yakın · {s['Uygunsuz']} uygunsuz · "
          f"{sum(x['sonuc'] == 'RET' for x in r['parca'])} / {len(r['parca'])} parça RET")
    for x in r["parca"]:
        if x["sonuc"] == "RET":
            print(f"[!] {x['parca']} · {x['numune']}: " + ", ".join(x["uygunsuz"]))
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
