"""
Kesim Kat Planı — Workers / Workless kod bloğu
Tekstil ve Konfeksiyon › Planlama › Planlama Uzmanı

Sipariş beden dağılımından pastal (marker) kombinasyonlarını ve kat adetlerini çıkarır:
  - Kısıtlar: azami kat sayısı (kumaş türü ve bıçak yüksekliği), pastala sığan azami ürün sayısı (masa boyu),
    asgari kat; hedef kesim = sipariş × (1 + fazla kesim %).
  - Yöntem: her adımda kalan adetlerden en çok ürünü kesen (kat × pastal oranı) pastal seçilir; oran, kalan
    adetlerle orantılı dağıtılır. Kalan küçük miktarlar, fazla kesimi en aza indiren tek bir son pastalla kapatılır.
  - Beden başına tüketim (m/adet) verilirse pastal boyu ve kumaş ihtiyacı (uç payları dahil) hesaplanır.
İnternete bağlanmaz.

Kullanım:
    python main.py                                         # örnek siparişle dener
    python main.py --siparis siparis.xlsx --max-kat 80 --max-urun 6 --fazla-kesim 2
    python main.py --siparis siparis.xlsx --tuketim tuketim.xlsx --uc-payi 3
"""
from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from collections import OrderedDict
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


def sayi(x, varsayilan=0.0):
    if x in (None, ""):
        return varsayilan
    if isinstance(x, (int, float)):
        return float(x)
    try:
        return float(str(x).strip().replace(".", "").replace(",", ".") if "," in str(x) else str(x).strip())
    except ValueError:
        return varsayilan


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


def siparis_oku(yol: Path) -> "OrderedDict[str, OrderedDict[str, int]]":
    """Renk → {beden: adet}. Uzun (Renk, Beden, Adet) veya geniş (Renk, S, M, L …) biçim."""
    s = tablo_oku(yol)
    kb = [katla(x) for x in s[0]]
    bul = lambda *a: next((kb.index(katla(x)) for x in a if katla(x) in kb), None)  # noqa: E731
    i_r, i_b, i_a = bul("renk", "renk kodu"), bul("beden"), bul("adet", "miktar", "sipariş adedi")
    i_m = bul("model", "artikel")
    sonuc: OrderedDict[str, OrderedDict[str, int]] = OrderedDict()
    if i_b is not None and i_a is not None:
        for r in s[1:]:
            renk = str(r[i_r] if i_r is not None else "-").strip() or "-"
            if i_m is not None and r[i_m]:
                renk = f"{str(r[i_m]).strip()} {renk}"
            sonuc.setdefault(renk, OrderedDict())[str(r[i_b]).strip()] = int(sayi(r[i_a]))
        return sonuc
    sabit = {i_r, i_m, bul("toplam"), bul("sipariş no", "po")}
    bedenler = [(i, str(b).strip()) for i, b in enumerate(s[0]) if i not in sabit and str(b or "").strip()]
    for r in s[1:]:
        renk = str(r[i_r] if i_r is not None else "-").strip() or "-"
        if i_m is not None and r[i_m]:
            renk = f"{str(r[i_m]).strip()} {renk}"
        sonuc[renk] = OrderedDict((b, int(sayi(r[i] if i < len(r) else 0))) for i, b in bedenler)
    return sonuc


def tuketim_oku(yol: Path | None) -> dict[str, float]:
    if not yol:
        return {}
    s = tablo_oku(yol)
    kb = [katla(x) for x in s[0]]
    if "beden" in kb:
        i_b = kb.index("beden")
        i_t = next(i for i, x in enumerate(kb) if x.startswith(("tuketim", "metre", "uzunluk")))
        return {str(r[i_b]).strip(): sayi(r[i_t]) for r in s[1:] if r[i_b]}
    return {str(b).strip(): sayi(v) for b, v in zip(s[0], s[1]) if str(b or "").strip() and sayi(v)}


# ----------------------------------------------------------------------------
# Planlama
# ----------------------------------------------------------------------------

def oran_sec(kalan: dict[str, int], kat: int, max_urun: int) -> dict[str, int]:
    """Bu kat sayısında kesilebilecek en çok ürünü veren pastal oranı (kalan adetle orantılı, ≤ max_urun)."""
    taban = {b: k // kat for b, k in kalan.items()}
    if sum(taban.values()) <= max_urun:
        return taban
    # max_urun'u kalan adetlerle orantılı dağıt (en büyük kalan), her beden en çok taban kadar
    toplam = sum(kalan.values())
    ham = {b: max_urun * k / toplam for b, k in kalan.items()}
    r = {b: min(taban[b], int(ham[b])) for b in kalan}
    while sum(r.values()) < max_urun:
        aday = [b for b in kalan if r[b] < taban[b]]
        if not aday:
            break
        b = max(aday, key=lambda b: (ham[b] - r[b], kalan[b]))
        r[b] += 1
    return r


def son_pastal(kalan: dict[str, int], max_kat: int, max_urun: int, min_kat: int = 1) -> tuple[int, int, dict[str, int]] | None:
    """Kalan adetleri tek pastalla kapat: (fazla kesim, kat, oran) — fazla kesimi en az, eşitlikte katı yüksek olan."""
    en_iyi = None
    for kat in range(min_kat, max_kat + 1):
        r = {b: math.ceil(k / kat) for b, k in kalan.items()}
        if sum(r.values()) > max_urun:
            continue
        fazla = sum(r[b] * kat - kalan[b] for b in kalan)
        if en_iyi is None or (fazla, -kat) < (en_iyi[0], -en_iyi[1]):
            en_iyi = (fazla, kat, r)
    return en_iyi


def iki_pastal(kalan: dict[str, int], max_kat: int, max_urun: int, min_kat: int):
    """Kalanı iki pastalla kapat: (fazla kesim, [(kat, oran), (kat, oran)])."""
    en_iyi = None
    for kat in range(max_kat, min_kat - 1, -1):
        r = oran_sec(kalan, kat, max_urun)
        if not sum(r.values()):
            continue
        geri = {b: k - r.get(b, 0) * kat for b, k in kalan.items()}
        geri = {b: k for b, k in geri.items() if k > 0}
        if not geri:
            return (0, [(kat, r)])
        s2 = son_pastal(geri, max_kat, max_urun, min_kat)
        if s2 and (en_iyi is None or s2[0] < en_iyi[0]):
            en_iyi = (s2[0], [(kat, r), (s2[1], s2[2])])
    return en_iyi


def planla(hedef: dict[str, int], max_kat: int, max_urun: int, min_kat: int) -> list[dict]:
    kalan = {b: a for b, a in hedef.items() if a > 0}
    sinir = max(max_urun, round(0.02 * sum(kalan.values())))     # kapanışta kabul edilen en çok fazla kesim (adet)
    pastallar = []

    def ekle(kat, oran, son=False):
        pastallar.append({"kat": kat, "oran": {b: v for b, v in oran.items() if v}, "son": son})

    while sum(kalan.values()) > 0:
        s1 = son_pastal(kalan, max_kat, max_urun, min_kat)
        if s1 and s1[0] <= sinir:
            ekle(s1[1], s1[2], True)
            break
        if sum(kalan.values()) <= 2 * max_urun * max_kat:
            s2 = iki_pastal(kalan, max_kat, max_urun, min_kat)
            if s2 and s2[0] <= sinir:
                for kat, oran in s2[1]:
                    ekle(kat, oran, True)
                break
        en_iyi = None
        for kat in range(max_kat, min_kat - 1, -1):
            r = oran_sec(kalan, kat, max_urun)
            kesim = kat * sum(r.values())
            if kesim and (en_iyi is None or kesim > en_iyi[0]):
                en_iyi = (kesim, kat, r)
        if en_iyi is None:
            # Asgari kat ile normal pastal kurulamıyor: en az fazla kesimli kapanışı kabul et
            secenek = [x for x in (son_pastal(kalan, max_kat, max_urun, min_kat), ) if x]
            s2 = iki_pastal(kalan, max_kat, max_urun, min_kat)
            if s2 and (not secenek or s2[0] < secenek[0][0]):
                for kat, oran in s2[1]:
                    ekle(kat, oran, True)
            elif secenek:
                ekle(secenek[0][1], secenek[0][2], True)
            else:
                s1 = son_pastal(kalan, max_kat, max_urun, 1)
                ekle(s1[1], s1[2], True)
            break
        _, kat, r = en_iyi
        ekle(kat, r)
        for b, v in r.items():
            kalan[b] -= v * kat
        kalan = {k: v for k, v in kalan.items() if v > 0}
    return pastallar


def calistir(siparis_yolu: Path, cikti: Path, max_kat: int = 80, max_urun: int = 6, min_kat: int = 5,
             fazla_kesim: float = 0.0, tuketim_yolu: Path | None = None, uc_payi_cm: float = 3.0) -> dict:
    if max_urun < 1 or max_kat < 1 or min_kat < 1 or min_kat > max_kat:
        raise SystemExit("Kısıtlar geçersiz: max-urun ≥ 1, 1 ≤ min-kat ≤ max-kat olmalı.")
    siparis = siparis_oku(siparis_yolu)
    tuketim = tuketim_oku(tuketim_yolu)
    sonuc = OrderedDict()
    no = 0
    for renk, bedenler in siparis.items():
        hedef = OrderedDict((b, math.ceil(a * (1 + fazla_kesim / 100))) for b, a in bedenler.items())
        pastallar = planla(hedef, max_kat, max_urun, min_kat)
        kesilen = OrderedDict((b, 0) for b in bedenler)
        for p in pastallar:
            no += 1
            p["no"] = no
            for b, v in p["oran"].items():
                kesilen[b] += v * p["kat"]
            if tuketim:
                p["boy"] = sum(tuketim.get(b, 0) * v for b, v in p["oran"].items()) + 2 * uc_payi_cm / 100
                p["kumas"] = p["boy"] * p["kat"]
        sonuc[renk] = {"siparis": bedenler, "hedef": hedef, "pastallar": pastallar, "kesilen": kesilen}
    _rapor(sonuc, max_kat, max_urun, min_kat, fazla_kesim, bool(tuketim), uc_payi_cm, cikti)
    return sonuc


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
SARI = PatternFill("solid", fgColor="FFF4CE")
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")


def _baslik(ws, satir=1):
    for c in ws[satir]:
        c.fill, c.font = BASLIK_DOLGU, BASLIK_YAZI


def oran_metni(oran: dict[str, int]) -> str:
    return " · ".join(f"{b}×{v}" for b, v in oran.items())


def _rapor(sonuc, max_kat, max_urun, min_kat, fazla_kesim, tuketimli, uc_payi, cikti):
    wb = Workbook()
    p = wb.active
    p.title = "Kat Planı"
    bas = ["Pastal No", "Renk", "Pastal Oranı", "Pastaldaki Ürün", "Kat Sayısı", "Kesilen Adet"]
    if tuketimli:
        bas += ["Pastal Boyu (m)", "Kumaş (m)"]
    p.append(bas)
    _baslik(p)
    for renk, x in sonuc.items():
        for pa in x["pastallar"]:
            satir = [pa["no"], renk, oran_metni(pa["oran"]), sum(pa["oran"].values()), pa["kat"], pa["kat"] * sum(pa["oran"].values())]
            if tuketimli:
                satir += [round(pa["boy"], 2), round(pa["kumas"], 1)]
            p.append(satir)
            if pa["kat"] < min_kat:
                p.cell(p.max_row, 5).fill = SARI
    for j, w in enumerate((10, 18, 34, 14, 10, 12, 14, 12), 1):
        p.column_dimensions[get_column_letter(j)].width = w
    p.freeze_panes = "C2"

    o = wb.create_sheet("Sipariş ve Kesim")
    o.append(["Renk", "Beden", "Sipariş", "Hedef (fazla kesim dahil)", "Kesilen", "Fark (kesilen − sipariş)", "Fark %"])
    _baslik(o)
    for renk, x in sonuc.items():
        for b, a in x["siparis"].items():
            k = x["kesilen"][b]
            o.append([renk, b, a, x["hedef"][b], k, k - a, (k - a) / a if a else None])
            o.cell(o.max_row, 7).number_format = "0.0%"
            if k < a:
                o.cell(o.max_row, 6).fill = KIRMIZI
            elif a and (k - a) / a > fazla_kesim / 100 + 0.05:
                o.cell(o.max_row, 6).fill = SARI
        top_s, top_k = sum(x["siparis"].values()), sum(x["kesilen"].values())
        o.append([renk, "TOPLAM", top_s, sum(x["hedef"].values()), top_k, top_k - top_s, (top_k - top_s) / top_s if top_s else None])
        o.cell(o.max_row, 7).number_format = "0.0%"
        for c in o[o.max_row]:
            c.font = Font(bold=True)
    for j, w in enumerate((18, 9, 10, 14, 10, 14, 9), 1):
        o.column_dimensions[get_column_letter(j)].width = w

    b = wb.create_sheet("Bilgi")
    pastal_say = sum(len(x["pastallar"]) for x in sonuc.values())
    satirlar = [["Kısıtlar", f"azami kat {max_kat}, pastalda azami ürün {max_urun}, asgari kat {min_kat}, fazla kesim %{fazla_kesim:g}"],
                ["Pastal sayısı", pastal_say],
                ["Yöntem", "her adımda kalan adetlerden en çok ürünü kesen (kat × oran) pastal seçilir; kalan küçük miktarlar fazla "
                           "kesimi en aza indiren tek bir son pastalla kapatılır. Sonuç en iyi (optimum) olmayabilir; pastal sayısını "
                           "azaltmak için azami ürün veya kat sayısını değiştirip karşılaştırın"],
                ["Sarı", "asgari kat altındaki pastal (son pastal) veya hedefin belirgin üzerinde kesim"],
                ["Kırmızı", "sipariş adedinin altında kesim"]]
    if tuketimli:
        satirlar.append(["Pastal boyu", f"Σ (oran × beden tüketimi) + 2 × {uc_payi:g} cm uç payı; kumaş = pastal boyu × kat"])
        satirlar.append(["Toplam kumaş (m)", round(sum(pa["kumas"] for x in sonuc.values() for pa in x["pastallar"]), 1)])
    for s in satirlar:
        b.append(s)
    b.column_dimensions["A"].width = 18
    b.column_dimensions["B"].width = 130
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ornek = BURASI / "ornek_veri"
    ap = argparse.ArgumentParser(description="Sipariş beden dağılımından pastal oranlarını ve kat adetlerini çıkarır.")
    ap.add_argument("--siparis", type=Path, default=ornek / "siparis.csv", help="Sipariş: Renk, S, M, L … (veya Renk, Beden, Adet)")
    ap.add_argument("--max-kat", type=int, default=80, help="Azami kat sayısı (varsayılan 80)")
    ap.add_argument("--max-urun", type=int, default=6, help="Bir pastala sığan azami ürün sayısı (varsayılan 6)")
    ap.add_argument("--min-kat", type=int, default=5, help="Asgari kat sayısı (varsayılan 5)")
    ap.add_argument("--fazla-kesim", type=float, default=0.0, help="Sipariş üzerine kesim payı, %% (varsayılan 0)")
    ap.add_argument("--tuketim", type=Path, help="Beden başına kumaş tüketimi (m/adet): Beden, Tüketim — veya S, M, L … tek satır")
    ap.add_argument("--uc-payi", type=float, default=3.0, help="Pastal başı ve sonu için kat başına uç payı, cm (varsayılan 3)")
    ap.add_argument("--cikti", type=Path, default=Path("cikti") / "kesim_kat_plani.xlsx")
    a = ap.parse_args(argv)
    if a.siparis == ornek / "siparis.csv":
        a.tuketim = a.tuketim or ornek / "tuketim.csv"
        if a.fazla_kesim == 0.0:
            a.fazla_kesim = 2.0
    s = calistir(a.siparis, a.cikti, a.max_kat, a.max_urun, a.min_kat, a.fazla_kesim, a.tuketim, a.uc_payi)
    for renk, x in s.items():
        top_s, top_k = sum(x["siparis"].values()), sum(x["kesilen"].values())
        print(f"[OK] {renk}: {len(x['pastallar'])} pastal · sipariş {top_s} · kesim {top_k} (+{top_k - top_s})")
        for p in x["pastallar"]:
            print(f"      #{p['no']}: {oran_metni(p['oran'])} × {p['kat']} kat")
    print(f"[OK] Rapor: {a.cikti.resolve()}")


if __name__ == "__main__":
    sys.exit(main())
