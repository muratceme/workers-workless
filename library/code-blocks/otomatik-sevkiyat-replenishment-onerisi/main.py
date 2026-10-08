"""
Otomatik Sevkiyat (Replenishment) Önerisi — Workers / Workless kod bloğu
Perakende › Ürün Planlama ve Alokasyon › Alokasyon Uzmanı

Mağaza satış hızı, mağaza stoğu ve hedef stok gününe göre depodan sevk edilecek miktarları hesaplar:
  - Günlük satış hızı: son N günün satışı ÷ N (stokta olmadığı günler çıkarılabilir).
  - Hedef stok = max(asgari teşhir, günlük satış × (hedef stok günü + sevk süresi)); ihtiyaç = hedef − (stok + yoldaki).
  - Paket/koli katına yuvarlama; depo stoğu yetmezse en az stok gününe sahip mağazalara öncelik (adil paylaştırma).
  - Uyarılar: hiç satmayan fazla stok (mağazalar arası transfer adayı), depoda olmayan ihtiyaç, kırılmış teşhir.
İnternete bağlanmaz.

Kullanım:
    python main.py                                         # örnek verilerle dener
    python main.py --magaza-stok magaza_stok.xlsx --satislar satislar.xlsx --depo-stok depo.xlsx --hedef-gun 14
    python main.py ... --satis-gunu 28 --sevk-suresi 2 --asgari 2 --paket paketler.xlsx
"""
from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from collections import OrderedDict, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent


def kucuk(s) -> str:
    return str(s or "").replace("İ", "i").replace("I", "ı").lower().strip()


def katla(s) -> str:
    s = kucuk(s)
    for a, b in zip("ıöüşğç", "iousgc"):
        s = s.replace(a, b)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def sayi(x) -> float:
    if x in (None, ""):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def tarih(x) -> date | None:
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    for b in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(x or "").strip()[:10], b).date()
        except ValueError:
            continue
    return None


def tablo_oku(yol: Path) -> list[list]:
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, data_only=True, read_only=True)
        try:
            satirlar = [list(r) for r in wb.active.iter_rows(values_only=True)]
        finally:
            wb.close()
    else:
        metin = None
        for kod in ("utf-8-sig", "cp1254"):
            try:
                metin = yol.read_text(encoding=kod)
                break
            except UnicodeDecodeError:
                continue
        ilk = "\n".join(metin.splitlines()[:10])
        satirlar = list(csv.reader(metin.splitlines(), delimiter=max(";\t,", key=ilk.count)))
    return [r for r in satirlar if any(c not in (None, "") for c in r)]


def _bul(baslik, *adlar):
    b = [katla(x) for x in baslik]
    return next((b.index(katla(a)) for a in adlar if katla(a) in b), None)


def _al(r, i):
    return r[i] if i is not None and i < len(r) else None


MAGAZA = ("mağaza", "mağaza kodu", "store")
URUN = ("ürün kodu", "stok kodu", "sku", "barkod", "ürün")


def magaza_stok_oku(yol: Path) -> tuple[dict[tuple[str, str], dict], dict[str, str]]:
    s = tablo_oku(yol)
    b = s[0]
    i_m, i_u = _bul(b, *MAGAZA), _bul(b, *URUN)
    i_s, i_y = _bul(b, "stok", "mevcut stok", "eldeki stok"), _bul(b, "yoldaki", "yolda", "transit", "sevk edilen")
    i_min, i_ad = _bul(b, "asgari teşhir", "min teşhir", "asgari", "min"), _bul(b, "ürün adı", "açıklama", "ad")
    if None in (i_m, i_u, i_s):
        raise SystemExit(f"Mağaza stok dosyasında Mağaza, Ürün Kodu ve Stok gerekli. Başlıklar: {b}")
    stok, adlar = {}, {}
    for r in s[1:]:
        if not _al(r, i_m) or not _al(r, i_u):
            continue
        k = (str(_al(r, i_m)).strip(), str(_al(r, i_u)).strip())
        stok[k] = {"stok": sayi(_al(r, i_s)), "yolda": sayi(_al(r, i_y)),
                   "asgari": sayi(_al(r, i_min)) if _al(r, i_min) not in (None, "") else None}
        if _al(r, i_ad):
            adlar[k[1]] = str(_al(r, i_ad)).strip()
    return stok, adlar


def satislar_oku(yol: Path, bitis: date | None, gun: int) -> tuple[dict[tuple[str, str], float], date, date]:
    """Son 'gun' günün satış adedi (mağaza, ürün). Dönüş: (adetler, başlangıç, bitiş)."""
    s = tablo_oku(yol)
    b = s[0]
    i_t, i_m, i_u = _bul(b, "tarih", "satış tarihi"), _bul(b, *MAGAZA), _bul(b, *URUN)
    i_a = _bul(b, "adet", "miktar", "satış adedi")
    if None in (i_t, i_m, i_u, i_a):
        raise SystemExit(f"Satış dosyasında Tarih, Mağaza, Ürün Kodu ve Adet gerekli. Başlıklar: {b}")
    satir = [(tarih(_al(r, i_t)), str(_al(r, i_m)).strip(), str(_al(r, i_u)).strip(), sayi(_al(r, i_a))) for r in s[1:]]
    satir = [x for x in satir if x[0]]
    son = bitis or max(x[0] for x in satir)
    bas = son - timedelta(days=gun - 1)
    adet: dict[tuple[str, str], float] = defaultdict(float)
    for t, m, u, a in satir:
        if bas <= t <= son:
            adet[(m, u)] += a
    return dict(adet), bas, son


def depo_oku(yol: Path) -> dict[str, float]:
    s = tablo_oku(yol)
    i_u, i_s = _bul(s[0], *URUN), _bul(s[0], "stok", "kullanılabilir stok", "depo stoğu", "miktar")
    if i_u is None or i_s is None:
        raise SystemExit(f"Depo stok dosyasında Ürün Kodu ve Stok gerekli. Başlıklar: {s[0]}")
    return {str(_al(r, i_u)).strip(): sayi(_al(r, i_s)) for r in s[1:] if _al(r, i_u)}


def paket_oku(yol: Path | None) -> dict[str, int]:
    if not yol:
        return {}
    s = tablo_oku(yol)
    i_u, i_p = _bul(s[0], *URUN), _bul(s[0], "paket", "koli içi", "paket miktarı", "sevk katı")
    return {str(_al(r, i_u)).strip(): int(sayi(_al(r, i_p))) for r in s[1:] if _al(r, i_u) and sayi(_al(r, i_p)) > 1}


# ----------------------------------------------------------------------------
# Hesap
# ----------------------------------------------------------------------------

def yukari_kat(x: float, kat: int) -> int:
    return int(math.ceil(x / kat) * kat) if kat > 1 else int(math.ceil(x))


def calistir(stok_yolu: Path, satis_yolu: Path, depo_yolu: Path, cikti: Path, hedef_gun: float = 14, sevk_suresi: float = 2,
             satis_gunu: int = 28, asgari: float = 2, paket_yolu: Path | None = None, bitis: date | None = None,
             fazla_gun: float = 60) -> dict:
    stok, adlar = magaza_stok_oku(stok_yolu)
    satis, bas, son = satislar_oku(satis_yolu, bitis, satis_gunu)
    depo = depo_oku(depo_yolu)
    paket = paket_oku(paket_yolu)
    satirlar = []
    for (m, u), x in sorted(stok.items()):
        hiz = satis.get((m, u), 0.0) / satis_gunu
        asg = x["asgari"] if x["asgari"] is not None else asgari
        hedef = max(asg, hiz * (hedef_gun + sevk_suresi))
        mevcut = x["stok"] + x["yolda"]
        ihtiyac = max(0.0, hedef - mevcut)
        oneri = yukari_kat(ihtiyac, paket.get(u, 1)) if ihtiyac > 0 else 0
        stok_gunu = mevcut / hiz if hiz else None
        satirlar.append({"magaza": m, "urun": u, "ad": adlar.get(u, ""), "stok": x["stok"], "yolda": x["yolda"], "satis": satis.get((m, u), 0.0),
                         "hiz": hiz, "stok_gunu": stok_gunu, "hedef": hedef, "ihtiyac": ihtiyac, "oneri": oneri, "sevk": 0, "asgari": asg,
                         "not": []})
    # Depo kısıtı: ürün bazında toplam öneri depo stoğunu aşarsa en az stok günü olanlara öncelik
    uyarilar = []
    urunler = defaultdict(list)
    for s in satirlar:
        urunler[s["urun"]].append(s)
    for u, ss in urunler.items():
        talep = sum(s["oneri"] for s in ss)
        kalan = depo.get(u, 0.0)
        if not talep:
            continue
        if kalan <= 0:
            for s in ss:
                if s["oneri"]:
                    s["not"].append("depoda stok yok")
            uyarilar.append(f"{u}: {talep:g} adet ihtiyaç var, depoda stok yok")
            continue
        kat = paket.get(u, 1)
        if talep <= kalan:
            for s in ss:
                s["sevk"] = s["oneri"]
            continue
        uyarilar.append(f"{u}: ihtiyaç {talep:g}, depo {kalan:g} — en az stok gününe sahip mağazalara öncelik verildi")
        # Tur tur paylaştır: her turda stok günü en düşük mağazaya bir paket
        while kalan >= kat:
            adaylar = [s for s in ss if s["sevk"] < s["oneri"]]
            if not adaylar:
                break
            s = min(adaylar, key=lambda s: ((s["stok"] + s["yolda"] + s["sevk"]) / s["hiz"] if s["hiz"] else float("inf"),
                                            -s["hiz"], s["magaza"]))
            s["sevk"] += kat
            kalan -= kat
        for s in ss:
            if s["sevk"] < s["oneri"]:
                s["not"].append(f"depo yetersiz: {s['oneri'] - s['sevk']:g} adet eksik")
    for s in satirlar:
        if s["hiz"] == 0 and s["stok"] > 0:
            s["not"].append("son dönemde satış yok")
        elif s["stok_gunu"] is not None and s["stok_gunu"] > fazla_gun:
            s["not"].append(f"fazla stok ({s['stok_gunu']:.0f} gün) — transfer adayı")
        if s["stok"] + s["yolda"] + s["sevk"] < s["asgari"]:
            s["not"].append("sevkten sonra da asgari teşhirin altında")
    _rapor(satirlar, depo, uyarilar, bas, son, hedef_gun, sevk_suresi, satis_gunu, asgari, fazla_gun, cikti)
    return {"satirlar": satirlar, "uyarilar": uyarilar, "donem": (bas, son)}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
SARI = PatternFill("solid", fgColor="FFF4CE")
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
MAVI = PatternFill("solid", fgColor="E8F0FE")


def _baslik(ws, satir=1):
    for c in ws[satir]:
        c.fill, c.font = BASLIK_DOLGU, BASLIK_YAZI


def _rapor(satirlar, depo, uyarilar, bas, son, hedef_gun, sevk_suresi, satis_gunu, asgari, fazla_gun, cikti):
    wb = Workbook()
    s = wb.active
    s.title = "Sevk Listesi"
    s.append(["Mağaza", "Ürün Kodu", "Ürün Adı", "Sevk Adedi"])
    _baslik(s)
    for x in sorted(satirlar, key=lambda x: (x["magaza"], x["urun"])):
        if x["sevk"]:
            s.append([x["magaza"], x["urun"], x["ad"], x["sevk"]])
    for j, w in enumerate((16, 14, 30, 11), 1):
        s.column_dimensions[get_column_letter(j)].width = w
    s.auto_filter.ref = s.dimensions

    d = wb.create_sheet("Hesap Detayı")
    d.append(["Mağaza", "Ürün Kodu", "Ürün Adı", "Stok", "Yoldaki", f"Satış (son {satis_gunu} gün)", "Günlük Satış", "Stok Günü",
              "Hedef Stok", "İhtiyaç", "Öneri (paket katı)", "Sevk", "Not"])
    _baslik(d)
    for x in sorted(satirlar, key=lambda x: (x["stok_gunu"] if x["stok_gunu"] is not None else 1e9, x["magaza"])):
        d.append([x["magaza"], x["urun"], x["ad"], x["stok"], x["yolda"] or None, x["satis"], round(x["hiz"], 2),
                  round(x["stok_gunu"], 1) if x["stok_gunu"] is not None else None, round(x["hedef"], 1), round(x["ihtiyac"], 1),
                  x["oneri"] or None, x["sevk"] or None, "; ".join(x["not"])])
        n = d.max_row
        if any("depo" in t for t in x["not"]):
            d.cell(n, 13).fill = KIRMIZI
        elif any("transfer" in t or "satış yok" in t for t in x["not"]):
            d.cell(n, 13).fill = MAVI
        elif x["not"]:
            d.cell(n, 13).fill = SARI
    for j, w in enumerate((16, 14, 28, 8, 8, 12, 10, 9, 10, 9, 11, 8, 50), 1):
        d.column_dimensions[get_column_letter(j)].width = w
    d.freeze_panes = "C2"
    d.auto_filter.ref = d.dimensions

    u = wb.create_sheet("Ürün Özeti")
    u.append(["Ürün Kodu", "Depo Stoğu", "Toplam İhtiyaç", "Toplam Sevk", "Depoda Kalan", "Mağaza Sayısı"])
    _baslik(u)
    gr = defaultdict(list)
    for x in satirlar:
        gr[x["urun"]].append(x)
    for k, xs in sorted(gr.items()):
        sevk = sum(x["sevk"] for x in xs)
        u.append([k, depo.get(k, 0), sum(x["oneri"] for x in xs), sevk, depo.get(k, 0) - sevk, sum(1 for x in xs if x["sevk"])])
    for j, w in enumerate((14, 12, 14, 12, 12, 12), 1):
        u.column_dimensions[get_column_letter(j)].width = w

    b = wb.create_sheet("Bilgi")
    for x in [["Satış dönemi", f"{bas:%d.%m.%Y} – {son:%d.%m.%Y} ({satis_gunu} gün)"],
              ["Hedef stok", f"max(asgari teşhir {asgari:g}, günlük satış × (hedef stok günü {hedef_gun:g} + sevk süresi {sevk_suresi:g}))"],
              ["İhtiyaç", "hedef stok − (mağaza stoğu + yoldaki); paket/koli katına yukarı yuvarlanır"],
              ["Depo kısıtı", "toplam ihtiyaç depo stoğunu aşarsa, her turda stok günü en düşük mağazaya bir paket verilir"],
              ["Fazla stok", f"stok günü {fazla_gun:g}'ın üstünde veya hiç satmayan ürünler — mağazalar arası transfer adayı"],
              ["Not", "Yeni açılan ürünlerde ve kampanya dönemlerinde geçmiş satış hızı yanıltıcı olabilir; bu kalemleri elle gözden geçirin."]]:
        b.append(x)
    for x in uyarilar:
        b.append(["Uyarı", x])
    b.column_dimensions["A"].width = 16
    b.column_dimensions["B"].width = 130
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ornek = BURASI / "ornek_veri"
    ap = argparse.ArgumentParser(description="Mağaza satış hızı ve hedef stok gününe göre depodan sevk önerisi hesaplar.")
    ap.add_argument("--magaza-stok", type=Path, default=ornek / "magaza_stok.csv", help="Mağaza, Ürün Kodu, Stok [, Yoldaki, Asgari Teşhir, Ürün Adı]")
    ap.add_argument("--satislar", type=Path, default=ornek / "satislar.csv", help="Tarih, Mağaza, Ürün Kodu, Adet")
    ap.add_argument("--depo-stok", type=Path, default=ornek / "depo_stok.csv", help="Ürün Kodu, Stok (kullanılabilir)")
    ap.add_argument("--paket", type=Path, help="Ürün Kodu, Paket (sevk katı)")
    ap.add_argument("--hedef-gun", type=float, default=14, help="Mağazada tutulacak hedef stok günü (varsayılan 14)")
    ap.add_argument("--sevk-suresi", type=float, default=2, help="Depodan mağazaya sevk süresi, gün (varsayılan 2)")
    ap.add_argument("--satis-gunu", type=int, default=28, help="Satış hızı için geriye bakılan gün (varsayılan 28)")
    ap.add_argument("--asgari", type=float, default=2, help="Mağazada asgari teşhir adedi (varsayılan 2)")
    ap.add_argument("--fazla-gun", type=float, default=60, help="Bu stok gününün üstü fazla stok sayılır (varsayılan 60)")
    ap.add_argument("--bitis", help="Satış döneminin son günü GG.AA.YYYY (varsayılan satışlardaki son tarih)")
    ap.add_argument("--cikti", type=Path, default=Path("cikti") / "sevk_onerisi.xlsx")
    a = ap.parse_args(argv)
    if a.magaza_stok == ornek / "magaza_stok.csv" and not a.paket:
        a.paket = ornek / "paketler.csv"
    s = calistir(a.magaza_stok, a.satislar, a.depo_stok, a.cikti, a.hedef_gun, a.sevk_suresi, a.satis_gunu, a.asgari, a.paket,
                 tarih(a.bitis) if a.bitis else None, a.fazla_gun)
    sevk = [x for x in s["satirlar"] if x["sevk"]]
    print(f"[OK] {len(s['satirlar'])} mağaza-ürün · {len(sevk)} sevk satırı · {sum(x['sevk'] for x in sevk)} adet · "
          f"satış dönemi {s['donem'][0]:%d.%m.%Y}–{s['donem'][1]:%d.%m.%Y}")
    for u in s["uyarilar"]:
        print(f"[!] {u}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")


if __name__ == "__main__":
    sys.exit(main())
