"""
Kümül Risk Kontrolü — Workers / Workless kod bloğu
Reasürans › Reasürans Uzmanı

Yürürlükteki poliçelerin deprem bedellerini kümül bölgesi ve lokasyon (aynı adres) bazında toplar; sizin verdiğiniz kapasitelerle karşılaştırır.
  Şirket payı bedeli = sigorta bedeli × şirket payı %            (koasürans payı; boşsa %100)
  Bölge kümülü       = bölgedeki yürürlükteki poliçelerin şirket payı bedelleri toplamı
  PML tutarı         = bölge kümülü × PML %                      (PML oranını siz verirsiniz; boşsa %100)
  Kullanım %         = PML tutarı ÷ bölge kapasitesi × 100;   ≥ 100 "Aşım",  ≥ --uyari (80) "Eşiğe yakın"
  Lokasyon kümülü    = aynı adres koduna (yoksa il + ilçe + adrese) sahip poliçelerin toplamı ÷ --lokasyon-kapasite
Yürürlükte = başlangıç ≤ rapor tarihi < bitiş. Bölge tanımları, PML oranları ve kapasiteler koda gömülü değildir.
Kod deprem modellemesi yapmaz; adresleri harita üzerinde eşleştirmez.
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


SUTUN = {"no": ("police no", "police numarasi"), "sigortali": ("sigortali", "sigortali adi"), "bolge": ("kumul bolgesi", "deprem bolgesi", "bolge", "cresta"),
         "il": ("il",), "ilce": ("ilce",), "adres": ("adres", "risk adresi"), "kod": ("adres kodu", "uavt", "uavt kodu", "lokasyon kodu"),
         "bedel": ("deprem bedeli", "sigorta bedeli", "bedel"), "pay": ("sirket payi", "sirket payi yuzde", "pay"), "bas": ("baslangic", "baslangic tarihi"), "bit": ("bitis", "bitis tarihi")}
KAPASITE_SUTUN = {"bolge": ("kumul bolgesi", "deprem bolgesi", "bolge", "cresta"), "kapasite": ("kapasite", "limit", "bolge kapasitesi"), "pml": ("pml", "pml yuzde", "pml orani")}
Z = Decimal(0)


def oku(yol: Path, rapor: date) -> tuple[list[dict], dict, list[dict]]:
    uy, l, gorulen = [], [], set()
    disarida = {"bitmis": 0, "baslamamis": 0}
    for r in kayitlar(yol, SUTUN, ("no", "bedel")):
        no = metin(r["no"])
        if not no:
            continue
        kim = f"{no} (satır {r['_satir']})"
        if no in gorulen:
            uy.append(uyari("Yüksek", "Mükerrer poliçe", kim, "Poliçe numarası tekrar ediyor; ilk satır kullanıldı (kümül çift sayılmasın diye)."))
            continue
        gorulen.add(no)
        bedel, pay = para(r["bedel"]), para(r["pay"])
        if bedel is None or bedel < 0:
            uy.append(uyari("Yüksek", "Bedel okunamadı", kim, f"'{metin(r['bedel'])}'. Poliçe kümüle alınmadı."))
            continue
        if metin(r["pay"]) and (pay is None or not 0 < pay <= 100):
            uy.append(uyari("Yüksek", "Şirket payı okunamadı", kim, f"'{metin(r['pay'])}'. Pay %100 sayıldı."))
            pay = None
        bas, bit = tarih(r["bas"]), tarih(r["bit"])
        if (metin(r["bas"]) and not bas) or (metin(r["bit"]) and not bit):
            uy.append(uyari("Orta", "Tarih okunamadı", kim, "Başlangıç veya bitiş tarihi okunamadı; poliçe yürürlükte sayıldı."))
        if bit and bit <= rapor:
            disarida["bitmis"] += 1
            continue
        if bas and bas > rapor:
            disarida["baslamamis"] += 1
            continue
        adres = " ".join(x for x in (metin(r["il"]), metin(r["ilce"]), metin(r["adres"])) if x)
        if not metin(r["bolge"]):
            uy.append(uyari("Orta", "Kümül bölgesi boş", kim, "Poliçe 'Belirtilmemiş' bölgesinde toplandı."))
        if not metin(r["kod"]) and not metin(r["adres"]):
            uy.append(uyari("Orta", "Adres bilgisi yok", kim, "Adres kodu ve adres boş; lokasyon kümülüne alınamadı."))
        l.append({"no": no, "sigortali": metin(r["sigortali"]), "bolge": metin(r["bolge"]) or "Belirtilmemiş", "il": metin(r["il"]), "ilce": metin(r["ilce"]), "adres": adres,
                  "lokasyon": ("K:" + katla(r["kod"])) if metin(r["kod"]) else ("A:" + katla(adres)) if metin(r["adres"]) else "", "bedel": bedel, "pay": pay or Decimal(100),
                  "net": yuvarla(bedel * (pay or Decimal(100)) / 100)})
    return l, disarida, uy


def kapasite_oku(yol: Path) -> tuple[dict, list[dict]]:
    kap, uy = {}, []
    for r in kayitlar(yol, KAPASITE_SUTUN, ("bolge", "kapasite")):
        b = metin(r["bolge"])
        if not b:
            continue
        k, pml = para(r["kapasite"]), para(r["pml"])
        if k is None or k <= 0:
            uy.append(uyari("Yüksek", "Kapasite okunamadı", b, f"{yol.name} satır {r['_satir']}: '{metin(r['kapasite'])}'. Bölge kapasitesiz sayıldı."))
            continue
        if metin(r["pml"]) and (pml is None or not 0 < pml <= 100):
            uy.append(uyari("Yüksek", "PML oranı okunamadı", b, f"{yol.name} satır {r['_satir']}: '{metin(r['pml'])}'. PML %100 sayıldı."))
            pml = None
        kap[katla(b)] = {"kapasite": k, "pml": pml or Decimal(100)}
    return kap, uy


def durum_bul(kullanim: Decimal | None, uyari_esik: Decimal) -> str:
    return "Kapasite tanımsız" if kullanim is None else "Aşım" if kullanim > 100 else "Eşiğe yakın" if kullanim >= uyari_esik else "Uygun"


def kumul(policeler: list[dict], kapasiteler: dict, lokasyon_kapasite: Decimal | None = None, uyari_esik: Decimal = Decimal(80)) -> dict:
    uy, bolge = [], {}
    for p in policeler:
        b = bolge.setdefault(katla(p["bolge"]), {"ad": p["bolge"], "police": 0, "bedel": Z})
        b["police"] += 1
        b["bedel"] += p["net"]
    for k, b in bolge.items():
        c = kapasiteler.get(k)
        b["pml_oran"], b["kapasite"] = (c["pml"], c["kapasite"]) if c else (None, None)
        b["pml"] = yuvarla(b["bedel"] * c["pml"] / 100) if c else None
        b["kullanim"] = yuvarla(b["pml"] / c["kapasite"] * 100) if c else None
        b["bos"] = c["kapasite"] - b["pml"] if c else None
        b["durum"] = durum_bul(b["kullanim"], uyari_esik)
        if not c:
            uy.append(uyari("Yüksek", "Bölge kapasitesi tanımsız", b["ad"], f"Kapasite dosyasında bu bölge yok; {b['police']} poliçe, toplam {b['bedel']:,.2f} bedel karşılaştırılamadı."))
    lok = {}
    for p in policeler:
        if p["lokasyon"]:
            x = lok.setdefault(p["lokasyon"], {"adres": p["adres"] or p["lokasyon"][2:], "bolge": p["bolge"], "policeler": [], "bedel": Z})
            x["policeler"].append(p)
            x["bedel"] += p["net"]
    lokasyonlar = []
    for x in lok.values():
        x["kullanim"] = yuvarla(x["bedel"] / lokasyon_kapasite * 100) if lokasyon_kapasite else None
        x["asim"] = max(x["bedel"] - lokasyon_kapasite, Z) if lokasyon_kapasite else None
        x["durum"] = "Birden çok poliçe" if x["kullanim"] is None else durum_bul(x["kullanim"], uyari_esik)
        if len(x["policeler"]) > 1 or x["durum"] in ("Aşım", "Eşiğe yakın"):
            lokasyonlar.append(x)
    sira = {"Aşım": 0, "Eşiğe yakın": 1, "Kapasite tanımsız": 2, "Uygun": 3, "Birden çok poliçe": 3}
    return {"bolge": sorted(bolge.values(), key=lambda b: (sira[b["durum"]], -(b["kullanim"] or Z))), "lokasyon": sorted(lokasyonlar, key=lambda x: (sira[x["durum"]], -x["bedel"])),
            "policeler": policeler, "toplam": sum((p["net"] for p in policeler), Z), "lokasyon_kapasite": lokasyon_kapasite, "uyarilar": uy}


def rapor_yaz(cikti: Path, r: dict, rapor: date, disarida: dict) -> None:
    wb = Workbook()
    renk = {"Aşım": KIRMIZI, "Eşiğe yakın": SARI, "Kapasite tanımsız": MAVI, "Uygun": YESIL}
    ws = wb.active
    ws.title = "Bölge Kümülü"
    ws.append([f"Rapor tarihi {rapor:%d.%m.%Y} · yürürlükteki poliçe {len(r['policeler'])} · süresi bitmiş {disarida['bitmis']} ve henüz başlamamış {disarida['baslamamis']} poliçe alınmadı"])
    ws["A1"].font = Font(bold=True)
    _baslik(ws, ["Kümül Bölgesi", "Poliçe", "Toplam Bedel (şirket payı)", "PML %", "PML Tutarı", "Bölge Kapasitesi", "Kullanım %", "Boş Kapasite", "Durum"], (26, 8, 22, 8, 20, 20, 11, 20, 18), "A3")
    for b in r["bolge"]:
        _satir(ws, [b["ad"], b["police"], b["bedel"], b["pml_oran"], b["pml"], b["kapasite"], b["kullanim"], b["bos"], b["durum"]], para_s=(3, 5, 6, 7, 8), dolgu={9: renk[b["durum"]]})
    _satir(ws, ["Toplam", len(r["policeler"]), r["toplam"]], para_s=(3,))
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    ws = wb.create_sheet("Lokasyon Kümülü")
    _baslik(ws, ["Lokasyon (adres)", "Kümül Bölgesi", "Poliçe Sayısı", "Toplam Bedel (şirket payı)", "Lokasyon Kapasitesi", "Kullanım %", "Aşım Tutarı", "Durum", "Poliçeler"],
            (48, 20, 9, 22, 20, 11, 20, 18, 60))
    for x in r["lokasyon"]:
        _satir(ws, [x["adres"], x["bolge"], len(x["policeler"]), x["bedel"], r["lokasyon_kapasite"], x["kullanim"], x["asim"] or None, x["durum"],
                    ", ".join(f"{p['no']} ({p['sigortali']})" if p["sigortali"] else p["no"] for p in x["policeler"])], para_s=(4, 5, 6, 7), dolgu={8: renk.get(x["durum"])}, sar=(1, 9))
    ws.append([])
    ws.append(["Birden çok poliçesi olan veya kapasite eşiğine ulaşan lokasyonlar listelenir. Adresler yalnız yazım olarak eşleştirilir; adres kodu kullanmanız önerilir."])
    ws = wb.create_sheet("Poliçeler")
    _baslik(ws, ["Poliçe No", "Sigortalı", "Kümül Bölgesi", "Adres", "Sigorta Bedeli", "Şirket Payı %", "Şirket Payı Bedeli"], (14, 30, 20, 50, 20, 10, 20))
    for p in r["policeler"]:
        _satir(ws, [p["no"], p["sigortali"], p["bolge"], p["adres"], p["bedel"], p["pay"], p["net"]], para_s=(5, 7))
    uyari_sayfasi(wb, sirala(r["uyarilar"]))
    kaydet(wb, cikti)


def calistir(police_yol: Path, kapasite_yol: Path, cikti: Path, rapor: date, lokasyon_kapasite: Decimal | None = None, uyari_esik: Decimal = Decimal(80)) -> dict:
    kapasiteler, uy1 = kapasite_oku(kapasite_yol)
    policeler, disarida, uy2 = oku(police_yol, rapor)
    if not policeler:
        raise ValueError(f"{police_yol.name}: rapor tarihinde yürürlükte poliçe yok.")
    r = kumul(policeler, kapasiteler, lokasyon_kapasite, uyari_esik)
    r["uyarilar"], r["disarida"] = sirala(uy1 + uy2 + r["uyarilar"]), disarida
    rapor_yaz(cikti, r, rapor, disarida)
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Deprem kümülünü bölge ve lokasyon bazında toplayıp kapasitelerle karşılaştırır.")
    p.add_argument("--policeler", type=Path, default=ORNEK / "policeler.csv", help="Poliçe No, Sigortalı, Kümül Bölgesi, İl, İlçe, Adres, Adres Kodu, Deprem Bedeli, Şirket Payı %%, Başlangıç, Bitiş")
    p.add_argument("--kapasiteler", type=Path, default=ORNEK / "kapasiteler.csv", help="Kümül Bölgesi, Kapasite, PML %%")
    p.add_argument("--tarih", help="Rapor tarihi GG.AA.YYYY (varsayılan bugün; örnek veride 01.10.2026)")
    p.add_argument("--lokasyon-kapasite", help="Tek lokasyon için kapasite tutarı (örnek veride 1.000.000.000)")
    p.add_argument("--uyari", default="80", help="Eşiğe yakın sayılacak kullanım %% (varsayılan 80)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "kumul_risk_kontrolu.xlsx")
    a = p.parse_args(argv)
    for y in (a.policeler, a.kapasiteler):
        if not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    ornek = a.policeler == ORNEK / "policeler.csv"
    rapor = tarih(a.tarih) if a.tarih else (date(2026, 10, 1) if ornek else date.today())
    lk = para(a.lokasyon_kapasite) if a.lokasyon_kapasite else (Decimal(1_000_000_000) if ornek else None)
    esik = para(a.uyari)
    if not rapor or esik is None or not 0 < esik <= 100 or (a.lokasyon_kapasite and (lk is None or lk <= 0)):
        print("[X] --tarih GG.AA.YYYY biçiminde; --uyari 0-100 arasında; --lokasyon-kapasite pozitif olmalı")
        return 2
    try:
        r = calistir(a.policeler, a.kapasiteler, a.cikti, rapor, lk, esik)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    print(f"[OK] {len(r['policeler'])} yürürlükteki poliçe · toplam bedel {r['toplam']:,.2f} · {len(r['bolge'])} bölge")
    for b in r["bolge"]:
        if b["durum"] != "Uygun":
            print(f"[!] Bölge {b['ad']}: {b['durum']}" + (f" · kullanım %{b['kullanim']}" if b["kullanim"] is not None else ""))
    for x in r["lokasyon"]:
        if x["durum"] == "Aşım":
            print(f"[!] Lokasyon {x['adres']}: {len(x['policeler'])} poliçe · {x['bedel']:,.2f} · aşım {x['asim']:,.2f}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
