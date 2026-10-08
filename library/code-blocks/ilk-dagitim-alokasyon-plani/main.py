"""
İlk Dağıtım (Alokasyon) Planı — Workers / Workless kod bloğu
Perakende › Ürün Planlama ve Alokasyon › Alokasyon Uzmanı

Yeni sezon ürünlerinin depo stoğunu mağazalara dağıtır:
  - Depo stoğunun bir kısmı tamamlama (replenishment) için depoda tutulur (rezerv %).
  - Mağaza ağırlığı = satış potansiyeli (geçmiş satış veya puan) × küme katsayısı (A/B/C).
  - Her model-renk-beden için dağıtılabilir stok, mağaza ağırlığı × mağazanın (veya kümenin) beden eğrisi
    oranında paylaştırılır; en büyük kalan yöntemiyle tam sayıya yuvarlanır, önce asgari teşhir adedi verilir.
  - İsteğe bağlı asorti (prepack) paketleriyle dağıtım; mağaza kapasitesi aşılırsa fazla depoya döner.
  - Kontroller: kırık seri (ana bedenlerden biri eksik giden mağaza), kapasite kullanımı, depoda kalan.
İnternete bağlanmaz.

Kullanım:
    python main.py                                          # örnek verilerle dener
    python main.py --stok depo_stok.xlsx --magazalar magazalar.xlsx --rezerv 30
    python main.py --stok depo_stok.xlsx --magazalar magazalar.xlsx --egri beden_egrisi.xlsx --asorti S=1,M=2,L=2,XL=1
"""
from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from collections import OrderedDict, defaultdict
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
    return re.sub(r"[^a-z0-9%]+", " ", s).strip()


def sayi(x, varsayilan: float | None = 0.0) -> float | None:
    if x in (None, ""):
        return varsayilan
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace("%", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
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


class Tablo:
    def __init__(self, yol: Path):
        s = tablo_oku(yol)
        self.ad, self.bas, self.satirlar = yol.name, s[0], s[1:]
        self.kb = [katla(x) for x in self.bas]

    def sutun(self, *adlar, zorunlu: str | None = None):
        i = next((self.kb.index(katla(a)) for a in adlar if katla(a) in self.kb), None)
        if i is None and zorunlu:
            raise SystemExit(f"{self.ad}: '{zorunlu}' sütunu bulunamadı. Başlıklar: {self.bas}")
        return i

    @staticmethod
    def al(r, i):
        v = r[i] if i is not None and i < len(r) else None
        return v.strip() if isinstance(v, str) else v


# ----------------------------------------------------------------------------
# Okuma
# ----------------------------------------------------------------------------

def stok_oku(yol: Path) -> tuple["OrderedDict[tuple[str, str], OrderedDict[str, int]]", dict]:
    """(model, renk) → {beden: adet}; uzun (Model, Renk, Beden, Stok) veya geniş (Model, Renk, S, M, L …) biçim."""
    t = Tablo(yol)
    i_m = t.sutun("model", "model kodu", "ürün", "ürün kodu", zorunlu="Model")
    i_r = t.sutun("renk", "renk kodu")
    i_b = t.sutun("beden")
    i_s = t.sutun("stok", "depo stoğu", "adet", "miktar")
    i_kat = t.sutun("kategori", "ürün grubu")
    stok: OrderedDict[tuple[str, str], OrderedDict[str, int]] = OrderedDict()
    kategori = {}
    if i_b is not None and i_s is not None:
        for r in t.satirlar:
            if t.al(r, i_m):
                k = (str(t.al(r, i_m)), str(t.al(r, i_r) or "-"))
                stok.setdefault(k, OrderedDict())[str(t.al(r, i_b))] = int(sayi(t.al(r, i_s)))
                if i_kat is not None:
                    kategori[k[0]] = str(t.al(r, i_kat) or "")
        return stok, kategori
    sabit = {i_m, i_r, i_kat, t.sutun("toplam")}
    bedenler = [(i, str(b).strip()) for i, b in enumerate(t.bas) if i not in sabit and str(b or "").strip()]
    for r in t.satirlar:
        if t.al(r, i_m):
            k = (str(t.al(r, i_m)), str(t.al(r, i_r) or "-"))
            stok[k] = OrderedDict((b, int(sayi(t.al(r, i)))) for i, b in bedenler)
            if i_kat is not None:
                kategori[k[0]] = str(t.al(r, i_kat) or "")
    return stok, kategori


def magazalar_oku(yol: Path) -> "OrderedDict[str, dict]":
    t = Tablo(yol)
    i_k = t.sutun("mağaza", "mağaza kodu", "mağaza adı", zorunlu="Mağaza")
    i_kume = t.sutun("küme", "grade", "sınıf", "mağaza sınıfı")
    i_p = t.sutun("potansiyel", "satış potansiyeli", "geçmiş satış", "satış adedi", "ağırlık", "puan")
    i_kap = t.sutun("kapasite", "kapasite adet", "teşhir kapasitesi")
    i_akt = t.sutun("aktif", "durum")
    m: OrderedDict[str, dict] = OrderedDict()
    for r in t.satirlar:
        if not t.al(r, i_k):
            continue
        if i_akt is not None and katla(t.al(r, i_akt)) in ("hayir", "pasif", "kapali", "0"):
            continue
        m[str(t.al(r, i_k))] = {"kume": str(t.al(r, i_kume) or "-").upper(), "potansiyel": sayi(t.al(r, i_p), 1.0),
                                "kapasite": sayi(t.al(r, i_kap), None)}
    if not m:
        raise SystemExit(f"{t.ad}: aktif mağaza yok.")
    return m


def egri_oku(yol: Path | None) -> dict[str, dict[str, float]]:
    """Anahtar: mağaza kodu veya küme ('A', 'B' …) veya '*' (genel). Değer: beden → pay."""
    if not yol:
        return {}
    t = Tablo(yol)
    i_k = t.sutun("mağaza", "küme", "mağaza küme", "kapsam", zorunlu="Mağaza/Küme")
    i_b, i_p = t.sutun("beden"), t.sutun("pay", "oran", "pay %", "beden payı")
    egri: dict[str, dict[str, float]] = defaultdict(dict)
    if i_b is not None and i_p is not None:
        for r in t.satirlar:
            if t.al(r, i_k) and t.al(r, i_b):
                egri[str(t.al(r, i_k)).upper() if len(str(t.al(r, i_k))) == 1 else str(t.al(r, i_k))][str(t.al(r, i_b))] = sayi(t.al(r, i_p))
    else:
        bedenler = [(i, str(b).strip()) for i, b in enumerate(t.bas) if i != i_k and str(b or "").strip()]
        for r in t.satirlar:
            if t.al(r, i_k):
                k = str(t.al(r, i_k))
                egri[k.upper() if len(k) == 1 else k] = {b: sayi(t.al(r, i)) for i, b in bedenler}
    return {k: {b: p / sum(v.values()) for b, p in v.items()} for k, v in egri.items() if sum(v.values()) > 0}


# ----------------------------------------------------------------------------
# Dağıtım
# ----------------------------------------------------------------------------

def en_buyuk_kalan(toplam: int, agirliklar: dict[str, float]) -> dict[str, int]:
    """toplam adedi ağırlıklara göre tam sayılara böler (Hamilton yöntemi); eşitlikte ağırlığı büyük olan önce."""
    s = sum(agirliklar.values())
    if toplam <= 0 or s <= 0:
        return {k: 0 for k in agirliklar}
    ham = {k: toplam * w / s for k, w in agirliklar.items()}
    tam = {k: int(math.floor(v)) for k, v in ham.items()}
    kalan = toplam - sum(tam.values())
    for k in sorted(ham, key=lambda k: (-(ham[k] - tam[k]), -agirliklar[k], k))[:kalan]:
        tam[k] += 1
    return tam


def beden_payi(magaza: str, kume: str, bedenler: list[str], egri: dict, stok_payi: dict[str, float]) -> dict[str, float]:
    """Mağazanın (yoksa kümesinin, yoksa genel '*') eğrisi, modelin bedenleri içinde yeniden normalleştirilir.
    Eğri modelin bedenlerini içermiyorsa (ör. pantolon numaraları) depo stoğunun beden dağılımı kullanılır."""
    for anahtar in (magaza, kume, "*"):
        e = egri.get(anahtar)
        if e:
            toplam = sum(e.get(b, 0.0) for b in bedenler)
            if toplam > 0:
                return {b: e.get(b, 0.0) / toplam for b in bedenler}
    return stok_payi


def dagit(stok, magazalar, egri, rezerv: float, asgari: int, kume_kat: dict[str, float],
          asorti: dict[str, int] | None) -> tuple[dict, list[str]]:
    agirlik = {m: v["potansiyel"] * kume_kat.get(v["kume"], 1.0) for m, v in magazalar.items()}
    plan: dict[tuple[str, str, str, str], int] = defaultdict(int)        # (mağaza, model, renk, beden) → adet
    uyarilar = []
    for (model, renk), bedenler in stok.items():
        toplam_stok = sum(bedenler.values())
        if toplam_stok <= 0:
            continue
        stok_payi = {b: a / toplam_stok for b, a in bedenler.items()}
        if asorti:
            paket_boyu = sum(asorti.values())
            eksik = [b for b in asorti if b not in bedenler]
            if eksik:
                uyarilar.append(f"{model} {renk}: asorti bedenleri ({', '.join(eksik)}) stokta yok; tekli dağıtıldı")
            else:
                paket = min(bedenler[b] // n for b, n in asorti.items() if n)
                dagitilacak = int(math.floor(paket * (1 - rezerv / 100)))
                pay = en_buyuk_kalan(dagitilacak, agirlik)
                for m, p in pay.items():
                    for b, n in asorti.items():
                        plan[(m, model, renk, b)] += p * n
                continue
        for b, adet in bedenler.items():
            dagitilacak = int(math.floor(adet * (1 - rezerv / 100)))
            talep = {m: agirlik[m] * beden_payi(m, magazalar[m]["kume"], list(bedenler), egri, stok_payi).get(b, 0.0)
                     for m in magazalar}
            talep = {m: w for m, w in talep.items() if w > 0}
            if not talep or dagitilacak <= 0:
                continue
            ilk = {}
            if asgari and dagitilacak >= asgari * len(talep):
                ilk = {m: asgari for m in talep}
            elif asgari:
                # yetmiyorsa asgari adet en yüksek ağırlıklı mağazalardan başlayarak verilir
                n = dagitilacak // asgari
                ilk = {m: asgari for m in sorted(talep, key=lambda m: (-talep[m], m))[:n]}
            kalan = dagitilacak - sum(ilk.values())
            ek = en_buyuk_kalan(kalan, talep)
            for m in talep:
                if ilk.get(m, 0) + ek.get(m, 0):
                    plan[(m, model, renk, b)] += ilk.get(m, 0) + ek.get(m, 0)
    # Kapasite: aşan mağazada en düşük ağırlıklı kalemlerden geri al (beden dengesini korumak için oransal azalt)
    for m, v in magazalar.items():
        if v["kapasite"] is None:
            continue
        anahtarlar = [k for k in plan if k[0] == m and plan[k]]
        toplam = sum(plan[k] for k in anahtarlar)
        if toplam > v["kapasite"]:
            yeni = en_buyuk_kalan(int(v["kapasite"]), {k: plan[k] for k in anahtarlar})
            for k in anahtarlar:
                plan[k] = yeni[k]
            uyarilar.append(f"{m}: kapasite {v['kapasite']:g} adet aşıldı ({toplam} adet); oransal azaltıldı, fazla depoda kaldı")
    return {k: v for k, v in plan.items() if v > 0}, uyarilar


def kirik_seriler(plan: dict, stok, ana_bedenler: list[str] | None) -> list[dict]:
    """Bir model-rengi alan mağazaya ana bedenlerden biri hiç gitmiyorsa kırık seri."""
    sonuc = []
    alan = defaultdict(set)
    for (m, model, renk, b), a in plan.items():
        alan[(m, model, renk)].add(b)
    for (m, model, renk), bedenler in sorted(alan.items()):
        mevcut = list(stok[(model, renk)])
        varsayilan = mevcut[1:-1] if len(mevcut) > 2 else mevcut          # en küçük ve en büyük beden hariç
        ana = [b for b in (ana_bedenler or varsayilan) if b in mevcut]
        eksik = [b for b in ana if b not in bedenler]
        if eksik:
            sonuc.append({"magaza": m, "model": model, "renk": renk, "eksik": eksik})
    return sonuc


def calistir(stok_yolu: Path, magaza_yolu: Path, cikti: Path, egri_yolu: Path | None = None, rezerv: float = 30.0,
             asgari: int = 1, kume_kat: dict[str, float] | None = None, asorti: dict[str, int] | None = None,
             ana_bedenler: list[str] | None = None) -> dict:
    stok, kategori = stok_oku(stok_yolu)
    magazalar = magazalar_oku(magaza_yolu)
    egri = egri_oku(egri_yolu)
    plan, uyarilar = dagit(stok, magazalar, egri, rezerv, asgari, kume_kat or {}, asorti)
    kirik = kirik_seriler(plan, stok, ana_bedenler)
    _rapor(stok, magazalar, plan, kirik, uyarilar, rezerv, cikti)
    return {"plan": plan, "kirik": kirik, "uyarilar": uyarilar, "stok": stok, "magazalar": magazalar}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
UYARI_DOLGU = PatternFill("solid", fgColor="FFF4CE")


def _baslik(ws, satir=1):
    for c in ws[satir]:
        c.fill, c.font = BASLIK_DOLGU, BASLIK_YAZI


def _rapor(stok, magazalar, plan, kirik, uyarilar, rezerv, cikti):
    wb = Workbook()
    o = wb.active
    o.title = "Model Özeti"
    o.append(["Model", "Renk", "Beden", "Depo Stoğu", "Dağıtılan", "Depoda Kalan", "Dağıtım %", "Mağaza Sayısı"])
    _baslik(o)
    for (model, renk), bedenler in stok.items():
        for b, a in bedenler.items():
            d = sum(v for k, v in plan.items() if k[1:] == (model, renk, b))
            n = len({k[0] for k in plan if k[1:] == (model, renk, b)})
            o.append([model, renk, b, a, d, a - d, d / a if a else None, n])
            o.cell(o.max_row, 7).number_format = "0%"
    toplam = sum(a for v in stok.values() for a in v.values())
    dagitilan = sum(plan.values())
    o.append(["TOPLAM", "", "", toplam, dagitilan, toplam - dagitilan, dagitilan / toplam if toplam else None, ""])
    o.cell(o.max_row, 7).number_format = "0%"
    for c in o[o.max_row]:
        c.font = Font(bold=True)
    for j, w in enumerate((14, 12, 8, 11, 11, 12, 10, 12), 1):
        o.column_dimensions[get_column_letter(j)].width = w
    o.freeze_panes = "D2"

    m = wb.create_sheet("Mağaza Özeti")
    m.append(["Mağaza", "Küme", "Potansiyel", "Kapasite", "Gönderilen Adet", "Kapasite Kullanımı", "Model-Renk Sayısı", "Kırık Seri"])
    _baslik(m)
    for k, v in magazalar.items():
        adet = sum(a for kk, a in plan.items() if kk[0] == k)
        mr = len({kk[1:3] for kk in plan if kk[0] == k})
        kr = sum(1 for x in kirik if x["magaza"] == k)
        m.append([k, v["kume"], v["potansiyel"], v["kapasite"], adet, adet / v["kapasite"] if v["kapasite"] else None, mr, kr or None])
        m.cell(m.max_row, 6).number_format = "0%"
        if kr:
            m.cell(m.max_row, 8).fill = UYARI_DOLGU
    for j, w in enumerate((16, 7, 11, 10, 14, 15, 15, 10), 1):
        m.column_dimensions[get_column_letter(j)].width = w

    x = wb.create_sheet("Dağıtım Matrisi")
    sutunlar = [(model, renk, b) for (model, renk), bedenler in stok.items() for b in bedenler]
    x.append(["Mağaza", "Küme"] + [f"{a} {r}" for a, r, _ in sutunlar] + ["Toplam"])
    x.append(["", ""] + [b for _, _, b in sutunlar] + [""])
    _baslik(x)
    _baslik(x, 2)
    for k, v in magazalar.items():
        satir = [plan.get((k, a, r, b), 0) or None for a, r, b in sutunlar]
        x.append([k, v["kume"]] + satir + [sum(s or 0 for s in satir)])
    x.append(["TOPLAM", ""] + [sum(plan.get((k, a, r, b), 0) for k in magazalar) for a, r, b in sutunlar] + [sum(plan.values())])
    for c in x[x.max_row]:
        c.font = Font(bold=True)
    x.column_dimensions["A"].width = 16
    for j in range(3, len(sutunlar) + 4):
        x.column_dimensions[get_column_letter(j)].width = 9
    x.freeze_panes = "C3"

    s = wb.create_sheet("Sevk Listesi")
    s.append(["Mağaza", "Model", "Renk", "Beden", "Adet"])
    _baslik(s)
    for (mg, model, renk, b), a in sorted(plan.items()):
        s.append([mg, model, renk, b, a])
    for j, w in enumerate((16, 14, 12, 8, 8), 1):
        s.column_dimensions[get_column_letter(j)].width = w
    s.auto_filter.ref = s.dimensions

    k = wb.create_sheet("Kontroller")
    k.append(["Kontrol", "Mağaza", "Model", "Renk", "Açıklama"])
    _baslik(k)
    for z in kirik:
        k.append(["Kırık seri", z["magaza"], z["model"], z["renk"], "Gitmeyen ana beden: " + ", ".join(z["eksik"])])
    for u in uyarilar:
        k.append(["Uyarı", "", "", "", u])
    k.append(["Bilgi", "", "", "", f"Tamamlama (replenishment) rezervi: her bedenin depo stoğunun %{rezerv:g} kadarı depoda tutuldu."])
    for j, w in enumerate((12, 16, 14, 12, 90), 1):
        k.column_dimensions[get_column_letter(j)].width = w
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ornek = BURASI / "ornek_veri"
    ap = argparse.ArgumentParser(description="Yeni sezon depo stoğunu mağaza potansiyeli ve beden eğrisine göre mağazalara dağıtır.")
    ap.add_argument("--stok", type=Path, default=ornek / "depo_stok.csv", help="Depo stoğu: Model, Renk, Beden, Stok — veya Model, Renk, S, M, L …")
    ap.add_argument("--magazalar", type=Path, default=ornek / "magazalar.csv", help="Mağazalar: Mağaza, Küme, Potansiyel [, Kapasite, Aktif]")
    ap.add_argument("--egri", type=Path, help="Beden eğrisi: Mağaza/Küme ('*' = genel), Beden, Pay — veya Küme, S, M, L …")
    ap.add_argument("--rezerv", type=float, default=30.0, help="Depoda tutulacak tamamlama rezervi, %% (varsayılan 30)")
    ap.add_argument("--asgari", type=int, default=1, help="Mağaza başına beden başına asgari teşhir adedi (varsayılan 1)")
    ap.add_argument("--kume", help="Küme katsayıları, ör. A=1,3 B=1 C=0,7 (potansiyel zaten ölçekliyse vermeyin)", nargs="*")
    ap.add_argument("--asorti", help="Asorti paketi, ör. S=1,M=2,L=2,XL=1 (verilirse paketle dağıtılır)")
    ap.add_argument("--ana-bedenler", help="Kırık seri kontrolündeki ana bedenler, ör. M,L (varsayılan: en küçük ve en büyük hariç)")
    ap.add_argument("--cikti", type=Path, default=Path("cikti") / "ilk_dagitim_plani.xlsx")
    a = ap.parse_args(argv)
    if a.stok == ornek / "depo_stok.csv" and not a.egri:
        a.egri = ornek / "beden_egrisi.csv"
    kume = {}
    for d in a.kume or []:
        k, _, v = d.partition("=")
        kume[k.strip().upper()] = sayi(v, 1.0)
    asorti = None
    if a.asorti:
        asorti = {}
        for parca in a.asorti.split(","):
            b, _, n = parca.partition("=")
            asorti[b.strip()] = int(sayi(n, 0))
    ana = [b.strip() for b in a.ana_bedenler.split(",")] if a.ana_bedenler else None
    s = calistir(a.stok, a.magazalar, a.cikti, a.egri, a.rezerv, a.asgari, kume, asorti, ana)
    toplam = sum(x for v in s["stok"].values() for x in v.values())
    print(f"[OK] {len(s['stok'])} model-renk · {len(s['magazalar'])} mağaza · depo {toplam} adet · dağıtılan {sum(s['plan'].values())} adet "
          f"· kırık seri {len(s['kirik'])}")
    for u in s["uyarilar"]:
        print(f"[!] {u}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")


if __name__ == "__main__":
    sys.exit(main())
