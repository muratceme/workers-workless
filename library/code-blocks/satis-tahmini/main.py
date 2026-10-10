"""
Satış Tahmini — Workers / Workless kod bloğu
Satış › Satış Müdürü

Geçmiş aylık satışlardan ürün grubu ve bölge bazında aylık tahmin üretir:
  - Seri: Ürün Grubu × Bölge; aylık toplam (eksik ay 0).
  - Yöntemler: naif (son ay), mevsimsel naif (geçen yılın aynı ayı), 3 aylık hareketli ortalama, Holt (doğrusal trend),
    Holt-Winters toplamsal ve çarpımsal (12 aylık mevsimsellik; en az 24 ay veri gerekir). Düzeltme katsayıları
    (α, β, γ) bir değer ızgarasında tek adımlı hata kareleri toplamını en küçükleyecek şekilde seçilir.
  - Seçim: son --test (6) ay dışarıda bırakılarak geriye dönük test yapılır; WAPE (Σ|hata| / Σ gerçek) en düşük yöntem
    seçilir, tüm veriyle yeniden kurulup --ufuk (12) ay tahmin edilir.
  - Yaklaşık aralık: tahmin ± 1,28 × test RMSE × √h (≈ %80); h ileri adım sayısı.
  - Uyarılar: WAPE > %30 (düşük güven), 24 aydan kısa seri (mevsimsellik yok), aykırı ay (geçen yılın aynı ayına oranı
    medyandan 4 MAD'den fazla sapan), sıfırı çok olan seri.
Rapor: tahmin tablosu, yöntem karşılaştırması, toplam (grafikli), mevsimsellik endeksi, geçmiş + tahmin, uyarılar.
İnternete bağlanmaz.

Kullanım:
    python main.py                                                 # örnek: 3 ürün grubu × 4 bölge, 36 ay
    python main.py --satislar satislar.xlsx --ufuk 12 --test 6 --deger Tutar
"""
from __future__ import annotations

import argparse
import csv
import itertools
import math
import re
import statistics
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
AYLAR = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"]
AY_KATLI = {"ocak": 1, "subat": 2, "mart": 3, "nisan": 4, "mayis": 5, "haziran": 6, "temmuz": 7, "agustos": 8, "eylul": 9, "ekim": 10, "kasim": 11, "aralik": 12}
M = 12
IZGARA = (0.1, 0.3, 0.5, 0.7, 0.9)
YONTEM_ADI = {"naif": "Naif (son ay)", "mevsimsel_naif": "Mevsimsel naif", "ho3": "Hareketli ortalama (3 ay)", "holt": "Holt (trend)",
              "hw_toplamsal": "Holt-Winters toplamsal", "hw_carpimsal": "Holt-Winters çarpımsal"}

SUTUNLAR = {"ay": ("ay", "donem", "tarih", "yil ay"), "grup": ("urun grubu", "grup", "kategori", "urun"), "bolge": ("bolge", "sube", "kanal"),
            "miktar": ("miktar", "adet", "satis miktari"), "tutar": ("tutar", "ciro", "satis tutari", "net satis")}


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(s or "").translate(_TR).lower()).split())


def ay_coz(x) -> tuple[int, int] | None:
    if isinstance(x, (date, datetime)):
        return x.year, x.month
    s = katla(x)
    for desen, sira in ((r"(\d{4}) (\d{1,2})(?: \d{1,2})?", (1, 2)), (r"(?:\d{1,2} )?(\d{1,2}) (\d{4})", (2, 1))):
        m = re.fullmatch(desen, s)
        if m:
            y, a = int(m[sira[0]]), int(m[sira[1]])
            return (y, a) if 1 <= a <= 12 else None
    m = re.fullmatch(r"([a-z]+) (\d{4})", s)
    if m and m[1] in AY_KATLI:
        return int(m[2]), AY_KATLI[m[1]]
    return None


def sayi(x) -> float | None:
    if x in (None, ""):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace("TL", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def metin(x) -> str:
    return str(x if x is not None else "").strip()


def ay_ileri(ym: tuple[int, int], n: int) -> tuple[int, int]:
    y, m = divmod(ym[1] - 1 + n, 12)
    return ym[0] + y, m + 1


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
    s = tablo_oku(yol)
    for bi, r in enumerate(s[:15]):
        b = [katla(c) for c in r]
        es = {}
        for alan, adlar in sozluk.items():
            j = next((b.index(a) for a in adlar if a in b and b.index(a) not in es.values()), None)
            if j is not None:
                es[alan] = j
        if all(z in es for z in zorunlu):
            return [{a: (r[j] if j < len(r) else None) for a, j in es.items()} | {"_satir": n} for n, r in enumerate(s[bi + 1:], bi + 2)]
    raise ValueError(f"{yol.name}: başlık satırı bulunamadı; gerekli sütunlar: {', '.join(zorunlu)}")


# ----------------------------------------------------------------------------
# Yöntemler: her biri (geçmiş, ufuk) → (tahminler, tek adımlı uydurma hataları)
# ----------------------------------------------------------------------------

def naif(y: list[float], h: int) -> list[float]:
    return [y[-1]] * h


def mevsimsel_naif(y: list[float], h: int) -> list[float]:
    return [y[len(y) - M + (i % M)] for i in range(h)]


def ho3(y: list[float], h: int) -> list[float]:
    return [sum(y[-3:]) / len(y[-3:])] * h


def holt(y: list[float], h: int) -> tuple[list[float], tuple]:
    en = None
    for a, b in itertools.product(IZGARA, IZGARA):
        l, t, sse = y[0], y[1] - y[0], 0.0
        for v in y[1:]:
            f = l + t
            sse += (v - f) ** 2
            l2 = a * v + (1 - a) * (l + t)
            t = b * (l2 - l) + (1 - b) * t
            l = l2
        if en is None or sse < en[0]:
            en = (sse, a, b, l, t)
    _, a, b, l, t = en
    return [l + (i + 1) * t for i in range(h)], (a, b)


def holt_winters(y: list[float], h: int, carpimsal: bool) -> tuple[list[float], tuple] | None:
    if len(y) < 2 * M or (carpimsal and min(y) <= 0):
        return None
    ilk, ikinci = y[:M], y[M:2 * M]
    l0 = sum(ilk) / M
    t0 = (sum(ikinci) - sum(ilk)) / (M * M)
    s0 = [v / l0 for v in ilk] if carpimsal else [v - l0 for v in ilk]
    en = None
    for a, b, g in itertools.product(IZGARA, IZGARA, IZGARA):
        l, t, s, sse = l0, t0, list(s0), 0.0
        for i in range(M, len(y)):
            v, si = y[i], s[i % M]
            f = (l + t) * si if carpimsal else l + t + si
            sse += (v - f) ** 2
            l2 = a * (v / si if carpimsal else v - si) + (1 - a) * (l + t)
            t = b * (l2 - l) + (1 - b) * t
            s[i % M] = g * (v / l2 if carpimsal else v - l2) + (1 - g) * si
            l = l2
        if en is None or sse < en[0]:
            en = (sse, a, b, g, l, t, list(s))
    _, a, b, g, l, t, s = en
    n = len(y)
    return [((l + (i + 1) * t) * s[(n + i) % M]) if carpimsal else (l + (i + 1) * t + s[(n + i) % M]) for i in range(h)], (a, b, g)


def tahmin_et(yontem: str, y: list[float], h: int) -> tuple[list[float], tuple] | None:
    if yontem == "naif":
        return naif(y, h), ()
    if yontem == "mevsimsel_naif":
        return (mevsimsel_naif(y, h), ()) if len(y) >= M else None
    if yontem == "ho3":
        return ho3(y, h), ()
    if yontem == "holt":
        return holt(y, h) if len(y) >= 3 else None
    return holt_winters(y, h, yontem == "hw_carpimsal")


def wape(gercek: list[float], tahmin: list[float]) -> float | None:
    payda = sum(abs(v) for v in gercek)
    return sum(abs(g - t) for g, t in zip(gercek, tahmin)) / payda if payda else None


# ----------------------------------------------------------------------------
# Seriler
# ----------------------------------------------------------------------------

@dataclass
class Seri:
    grup: str
    bolge: str
    aylar: list[tuple[int, int]]
    y: list[float]
    sonuclar: dict = field(default_factory=dict)
    secilen: str = ""
    tahmin: list[float] = field(default_factory=list)
    aralik: float = 0.0
    parametre: tuple = ()
    notlar: list[str] = field(default_factory=list)

    @property
    def ad(self) -> str:
        return f"{self.grup} · {self.bolge}"


def oku(yol: Path, deger: str) -> tuple[list[Seri], list[dict]]:
    uy = []
    toplam = defaultdict(lambda: defaultdict(float))
    alan = "tutar" if katla(deger) == "tutar" else "miktar"
    satirlar = kayitlar(yol, SUTUNLAR, ("ay", alan))
    for r in satirlar:
        ym, v = ay_coz(r.get("ay")), sayi(r.get(alan))
        if not ym or v is None:
            if metin(r.get("ay")) or metin(r.get(alan)):
                uy.append({"onem": "Orta", "tur": "Okunamayan satır", "kim": f"satır {r['_satir']}", "aciklama": "Ay veya değer okunamadı"})
            continue
        toplam[(metin(r.get("grup")) or "Tümü", metin(r.get("bolge")) or "Tümü")][ym] += v
    seriler = []
    for (g, b), d in sorted(toplam.items(), key=lambda i: (katla(i[0][0]), katla(i[0][1]))):
        bas, son = min(d), max(d)
        aylar, ym = [], bas
        while ym <= son:
            aylar.append(ym)
            ym = ay_ileri(ym, 1)
        seriler.append(Seri(g, b, aylar, [d.get(a, 0.0) for a in aylar]))
    return seriler, uy


def seri_isle(s: Seri, ufuk: int, test: int) -> list[dict]:
    uy = []
    n = len(s.y)
    if n < test + 6:
        s.notlar.append(f"{n} ay veri; geriye dönük test için yetersiz, naif yöntem kullanıldı")
        s.secilen, s.tahmin = "naif", naif(s.y, ufuk)
        uy.append({"onem": "Orta", "tur": "Kısa seri", "kim": s.ad, "aciklama": s.notlar[-1]})
        return uy
    egitim, gercek = s.y[:-test], s.y[-test:]
    for yontem in YONTEM_ADI:
        r = tahmin_et(yontem, egitim, test)
        if r is None:
            continue
        t = r[0]
        hatalar = [g - x for g, x in zip(gercek, t)]
        s.sonuclar[yontem] = {"wape": wape(gercek, t), "rmse": math.sqrt(sum(e * e for e in hatalar) / len(hatalar)), "test": t}
    gecerli = {k: v for k, v in s.sonuclar.items() if v["wape"] is not None}
    s.secilen = min(gecerli, key=lambda k: (gecerli[k]["wape"], list(YONTEM_ADI).index(k))) if gecerli else "naif"
    son = tahmin_et(s.secilen, s.y, ufuk)
    s.tahmin, s.parametre = [max(0.0, v) for v in son[0]], son[1]
    s.aralik = 1.28 * s.sonuclar.get(s.secilen, {"rmse": 0})["rmse"]
    if n < 2 * M:
        s.notlar.append(f"{n} ay veri (< 24): mevsimsel Holt-Winters denenmedi")
        uy.append({"onem": "Bilgi", "tur": "Mevsimsellik yok", "kim": s.ad, "aciklama": s.notlar[-1]})
    w = gecerli.get(s.secilen, {}).get("wape")
    if w is not None and w > 0.30:
        uy.append({"onem": "Orta", "tur": "Düşük güven", "kim": s.ad, "aciklama": f"En iyi yöntem ({YONTEM_ADI[s.secilen]}) testte WAPE %{w * 100:.0f}"})
    if sum(v == 0 for v in s.y) / n > 0.3:
        uy.append({"onem": "Bilgi", "tur": "Seyrek satış", "kim": s.ad, "aciklama": "Ayların %30'undan fazlasında satış yok; aylık tahmin yerine dönemsel toplam kullanın"})
    # Aykırı ay: geçen yılın aynı ayına oran (log); medyan ± 4 × MAD (sağlam ölçü). Önceki yılı aykırı olan ayın
    # oranı da bozulacağından (yankı) o ay atlanır.
    if n >= M + 6:
        idx = [i for i in range(M, n) if s.y[i] > 0 and s.y[i - M] > 0]
        oranlar = {i: math.log(s.y[i] / s.y[i - M]) for i in idx}
        if len(oranlar) >= 6:
            med = statistics.median(oranlar.values())
            mad = statistics.median(abs(v - med) for v in oranlar.values()) * 1.4826
            aykiri = set()
            for i in sorted(oranlar):
                if mad and abs(oranlar[i] - med) > 4 * mad and (i - M) not in aykiri:
                    aykiri.add(i)
                    y_, a = s.aylar[i]
                    uy.append({"onem": "Orta", "tur": "Aykırı ay", "kim": s.ad,
                               "aciklama": f"{AYLAR[a - 1]} {y_}: {s.y[i]:,.0f} (geçen yıl {s.y[i - M]:,.0f}); tek seferlik bir olaysa veriyi düzeltip "
                                           "yeniden çalıştırın".replace(",", ".")})
    return uy


def mevsim_endeksi(seriler: list[Seri]) -> dict[str, list[float | None]]:
    """Grup bazında ay endeksi: ayın ortalaması / yıllık ortalama × 100 (tam yıllar)."""
    sonuc = {}
    for g in sorted({s.grup for s in seriler}, key=katla):
        d = defaultdict(float)
        for s in seriler:
            if s.grup == g:
                for a, v in zip(s.aylar, s.y):
                    d[a] += v
        aylar = defaultdict(list)
        for (y, m), v in d.items():
            aylar[m].append(v)
        ort = statistics.fmean(v for lst in aylar.values() for v in lst) if aylar else 0
        sonuc[g] = [statistics.fmean(aylar[m]) / ort * 100 if aylar.get(m) and ort else None for m in range(1, 13)]
    return sonuc


def analiz_et(seriler: list[Seri], ufuk: int = 12, test: int = 6) -> dict:
    uy = []
    for s in seriler:
        uy += seri_isle(s, ufuk, test)
    son_ay = max(s.aylar[-1] for s in seriler)
    gelecek = [ay_ileri(son_ay, i + 1) for i in range(ufuk)]
    for s in seriler:
        if s.aylar[-1] < son_ay:
            kayma = (son_ay[0] - s.aylar[-1][0]) * 12 + son_ay[1] - s.aylar[-1][1]
            s.notlar.append(f"Son veri {AYLAR[s.aylar[-1][1] - 1]} {s.aylar[-1][0]}; tahmin {kayma} ay kaydırıldı")
            uy.append({"onem": "Orta", "tur": "Eksik son ay", "kim": s.ad, "aciklama": s.notlar[-1] + " (son aylarda satış yoksa 0 girin)"})
            uzun = tahmin_et(s.secilen, s.y, ufuk + kayma)
            s.tahmin = [max(0.0, v) for v in uzun[0][kayma:]] if uzun else s.tahmin
    tum_aylar = sorted({a for s in seriler for a in s.aylar})
    gecmis_toplam = [sum(s.y[s.aylar.index(a)] for s in seriler if a in s.aylar) for a in tum_aylar]
    gelecek_toplam = [sum(s.tahmin[i] for s in seriler) for i in range(ufuk)]
    son12 = sum(gecmis_toplam[-12:])
    once12 = sum(gecmis_toplam[-24:-12]) if len(gecmis_toplam) >= 24 else None
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    uy.sort(key=lambda u: (sira[u["onem"]], u["tur"], u["kim"]))
    return {"seriler": seriler, "gelecek": gelecek, "uyarilar": uy, "tum_aylar": tum_aylar, "gecmis_toplam": gecmis_toplam, "gelecek_toplam": gelecek_toplam,
            "son12": son12, "once12": once12, "gelecek12": sum(gelecek_toplam[:12]), "mevsim": mevsim_endeksi(seriler), "ufuk": ufuk, "test": test}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
RENK = {"Yüksek": "FDE2E1", "Orta": "FFF4CE", "Bilgi": "E8F0FE"}
UST = Alignment(vertical="top", wrap_text=True)
SF = "#,##0"


def _baslik(ws, basliklar, genislik=()):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "C2"


def ay_adi(a: tuple[int, int]) -> str:
    return f"{AYLAR[a[1] - 1]} {a[0]}"


def rapor_yaz(cikti: Path, s: dict, deger: str) -> None:
    wb = Workbook()
    th = wb.active
    th.title = "Tahmin"
    gel = s["gelecek"]
    _baslik(th, ["Ürün Grubu", "Bölge", "Yöntem"] + [ay_adi(a) for a in gel] + ["Toplam", "± Aralık (aylık, ≈%80)", "Test WAPE", "Not"],
            [16, 14, 20] + [10] * len(gel) + [12, 12, 9, 40])
    th.freeze_panes = "D2"
    for x in s["seriler"]:
        w = x.sonuclar.get(x.secilen, {}).get("wape")
        th.append([x.grup, x.bolge, YONTEM_ADI[x.secilen]] + [round(v) for v in x.tahmin] + [round(sum(x.tahmin)), round(x.aralik), w, "; ".join(x.notlar)])
        for j in range(4, len(gel) + 6):
            th.cell(th.max_row, j).number_format = SF
        th.cell(th.max_row, len(gel) + 6).number_format = "0.0%"
        if w is not None and w > 0.30:
            th.cell(th.max_row, len(gel) + 6).fill = PatternFill("solid", fgColor="FFF4CE")
    th.append(["Toplam", "", ""] + [round(v) for v in s["gelecek_toplam"]] + [round(sum(s["gelecek_toplam"]))])
    for j in range(1, len(gel) + 5):
        th.cell(th.max_row, j).font = Font(bold=True)
        th.cell(th.max_row, j).number_format = SF

    yk = wb.create_sheet("Yöntem Karşılaştırma")
    _baslik(yk, ["Ürün Grubu", "Bölge"] + [YONTEM_ADI[k] for k in YONTEM_ADI] + ["Seçilen", "Parametreler (α, β, γ)"], [16, 14] + [12] * len(YONTEM_ADI) + [20, 18])
    for x in s["seriler"]:
        yk.append([x.grup, x.bolge] + [x.sonuclar[k]["wape"] if k in x.sonuclar else None for k in YONTEM_ADI] + [YONTEM_ADI[x.secilen],
                                                                                                            ", ".join(f"{p:g}" for p in x.parametre)])
        for j, k in enumerate(YONTEM_ADI, 3):
            h = yk.cell(yk.max_row, j)
            h.number_format = "0.0%"
            if k == x.secilen:
                h.fill = PatternFill("solid", fgColor="E3F4E1")
    yk.append([])
    yk.append([f"Değerler son {s['test']} ayın geriye dönük test WAPE'sidir (Σ|gerçek − tahmin| / Σ gerçek). Boş: veri yetersiz."])

    tp = wb.create_sheet("Toplam")
    _baslik(tp, ["Ay", "Gerçekleşen", "Tahmin"], (12, 14, 14))
    tp.freeze_panes = "A2"
    for a, v in zip(s["tum_aylar"], s["gecmis_toplam"]):
        tp.append([ay_adi(a), round(v), None])
    for a, v in zip(gel, s["gelecek_toplam"]):
        tp.append([ay_adi(a), None, round(v)])
    for r in range(2, tp.max_row + 1):
        tp.cell(r, 2).number_format = tp.cell(r, 3).number_format = SF
    g = LineChart()
    g.title, g.height, g.width = f"Toplam {deger.lower()}: gerçekleşen ve tahmin", 9, 24
    g.add_data(Reference(tp, min_col=2, max_col=3, min_row=1, max_row=tp.max_row), titles_from_data=True)
    g.set_categories(Reference(tp, min_col=1, min_row=2, max_row=tp.max_row))
    tp.add_chart(g, "E2")
    tp.append([])
    tp.append(["Son 12 ay", round(s["son12"])])
    if s["once12"]:
        tp.append(["Önceki 12 ay", round(s["once12"]), None])
        tp.append(["Son 12 ay değişim", s["son12"] / s["once12"] - 1])
        tp.cell(tp.max_row, 2).number_format = "0.0%"
    tp.append(["Gelecek 12 ay (tahmin)", round(s["gelecek12"])])
    tp.append(["Tahmini değişim (son 12 aya göre)", s["gelecek12"] / s["son12"] - 1 if s["son12"] else None])
    tp.cell(tp.max_row, 2).number_format = "0.0%"

    me = wb.create_sheet("Mevsimsellik")
    _baslik(me, ["Ürün Grubu"] + AYLAR, [18] + [7] * 12)
    for gr, lst in s["mevsim"].items():
        me.append([gr] + [round(v) if v is not None else None for v in lst])
        for j, v in enumerate(lst, 2):
            if v is not None:
                me.cell(me.max_row, j).fill = PatternFill("solid", fgColor="E3F4E1" if v >= 115 else "FDE2E1" if v <= 85 else "FFFFFF")
    me.append([])
    me.append(["Ay endeksi = o ayın ortalama satışı / tüm ayların ortalaması × 100 (tüm bölgeler toplamı)."])

    gt = wb.create_sheet("Geçmiş + Tahmin")
    _baslik(gt, ["Ürün Grubu", "Bölge", "Ay", "Gerçekleşen", "Tahmin", "Alt (≈%80)", "Üst (≈%80)"], (16, 14, 10, 12, 12, 12, 12))
    for x in s["seriler"]:
        for a, v in zip(x.aylar, x.y):
            gt.append([x.grup, x.bolge, ay_adi(a), round(v), None, None, None])
        for i, (a, v) in enumerate(zip(gel, x.tahmin), 1):
            p = x.aralik * math.sqrt(i)
            gt.append([x.grup, x.bolge, ay_adi(a), None, round(v), round(max(0, v - p)), round(v + p)])
    gt.auto_filter.ref = f"A1:G{gt.max_row}"

    uy = wb.create_sheet("Uyarılar")
    _baslik(uy, ["Önem", "Tür", "Seri", "Açıklama"], (9, 20, 26, 100))
    uy.freeze_panes = "A2"
    for u in s["uyarilar"]:
        uy.append([u["onem"], u["tur"], u["kim"], u["aciklama"]])
        uy.cell(uy.max_row, 1).fill = PatternFill("solid", fgColor=RENK[u["onem"]])
        uy.cell(uy.max_row, 4).alignment = UST
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(yol: Path, cikti: Path, ufuk: int = 12, test: int = 6, deger: str = "Miktar") -> dict:
    seriler, uy = oku(yol, deger)
    if not seriler:
        raise ValueError(f"{yol.name}: satış bulunamadı")
    s = analiz_et(seriler, ufuk, test)
    s["uyarilar"] = uy + s["uyarilar"]
    rapor_yaz(cikti, s, deger)
    return s


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Geçmiş satışlar ve mevsimsellikten ürün grubu ve bölge bazında aylık satış tahmini üretir.")
    p.add_argument("--satislar", type=Path, default=ORNEK / "satislar.csv", help="Ay (2026-09 / 09.2026 / Eylül 2026), Ürün Grubu, Bölge, Miktar, Tutar")
    p.add_argument("--deger", choices=["Miktar", "Tutar"], default="Miktar", help="Hangi değer tahmin edilsin (varsayılan Miktar)")
    p.add_argument("--ufuk", type=int, default=12, help="Kaç ay ileri tahmin (varsayılan 12)")
    p.add_argument("--test", type=int, default=6, help="Geriye dönük test için ayrılan son ay sayısı (varsayılan 6)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "satis_tahmini.xlsx")
    a = p.parse_args(argv)
    if not a.satislar.exists():
        print(f"[X] Dosya bulunamadı: {a.satislar}")
        return 1
    if not (1 <= a.ufuk <= 36 and 1 <= a.test <= 12):
        print("[X] --ufuk 1–36, --test 1–12 olmalı")
        return 2
    try:
        s = calistir(a.satislar, a.cikti, a.ufuk, a.test, a.deger)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    print(f"[OK] {len(s['seriler'])} seri · son 12 ay {s['son12']:,.0f} · gelecek 12 ay tahmini {s['gelecek12']:,.0f}".replace(",", "."))
    for x in s["seriler"]:
        w = x.sonuclar.get(x.secilen, {}).get("wape")
        print(f"     {x.ad:<28} {YONTEM_ADI[x.secilen]:<26} WAPE {'—' if w is None else f'%{w * 100:.0f}'}")
    for u in s["uyarilar"]:
        if u["onem"] != "Bilgi":
            print(f"[!] {u['kim']} · {u['tur']}: {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
