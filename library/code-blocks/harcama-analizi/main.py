"""
Harcama Analizi — Workers / Workless kod bloğu
Satın Alma › Satın Alma Müdürü

Satın alma verisini kategori ve tedarikçi bazında ABC sınıflandırmasıyla analiz eder ve pazarlık önceliklerini çıkarır:
  - TL karşılığı: tutar × kur (döviz satırında kur yoksa satır analize alınmaz). Tutar yoksa miktar × birim fiyat.
  - Tedarikçi adları birleştirilir: büyük-küçük harf, Türkçe karakter ve şirket türü ekleri (A.Ş., Ltd. Şti., San. Tic.)
    yok sayılır; en sık yazılış gösterilir.
  - ABC: harcamaya göre azalan sırada; kendinden önceki kümülatif pay %80'in altındaysa A, %95'in altındaysa B, değilse C.
  - Kategori × tedarikçi: tedarikçi sayısı, en büyük tedarikçinin payı (tek kaynak bağımlılığı), C sınıfı (kuyruk)
    tedarikçilerin sayısı (dağınık harcama).
  - Fiyat farkı: aynı malzeme + birim aynı ayda (--fiyat-donem ceyrek ile çeyrekte) farklı fiyatlarla alındıysa ağırlıklı
    ortalamanın üzerindeki alımların fazlası ve en düşük fiyata göre fark. Dönem içinde karşılaştırmak enflasyon etkisini
    azaltır.
  - Fiyat artışı: malzemenin ilk ve son ay ağırlıklı ortalama fiyatı; tüm malzemelerin medyan artışını --artis puan aşan
    artışlar işaretlenir (enflasyon ortamında mutlak artış yanıltıcıdır).
  - Pazarlık öncelikleri kurallarla üretilir (README'de tablo): fiyat birliği, tek kaynak, tedarikçi konsolidasyonu,
    fiyat artışı.
Rapor: özet, kategori ABC, tedarikçi ABC, kategori × tedarikçi, fiyat farkları, fiyat artışları, pazarlık öncelikleri,
aylık trend, veri, uyarılar. İnternete bağlanmaz.

Kullanım:
    python main.py                                      # örnek: Ocak–Eylül 2026 satın almaları
    python main.py --veri satinalma.xlsx --a-esik 80 --b-esik 95 --artis 20
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
K2 = Decimal("0.01")
SIFIR = Decimal(0)
AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
SIRKET_EKLERI = {"a", "s", "as", "anonim", "sirketi", "sirket", "ltd", "limited", "sti", "san", "sanayi", "tic", "ticaret", "ve", "ltd sti", "koll", "kollektif"}

SUTUNLAR = {"tarih": ("tarih", "fatura tarihi", "siparis tarihi"), "tedarikci": ("tedarikci", "satici", "firma", "cari"),
            "kategori": ("kategori", "malzeme grubu", "harcama kategorisi"), "malzeme": ("malzeme", "malzeme adi", "kalem", "urun hizmet", "aciklama"),
            "miktar": ("miktar",), "birim": ("birim", "olcu birimi"), "fiyat": ("birim fiyat", "fiyat"), "tutar": ("tutar", "kdv haric tutar", "net tutar"),
            "doviz": ("doviz", "para birimi"), "kur": ("kur",), "departman": ("departman", "maliyet merkezi", "talep eden birim")}


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
    if x in (None, ""):
        return None
    if isinstance(x, (int, float)):
        return Decimal(str(x))
    s = str(x).strip().replace("TL", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def tl(x: Decimal | None) -> str:
    return "—" if x is None else f"{x.quantize(K2, ROUND_HALF_UP):,.0f}".replace(",", ".")


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


def tedarikci_anahtari(ad: str) -> str:
    k = [w for w in katla(ad).split() if w not in SIRKET_EKLERI]
    return " ".join(k) or katla(ad)


def ceyrek(d: date) -> str:
    return f"{d.year}-Ç{(d.month - 1) // 3 + 1}"


# ----------------------------------------------------------------------------
# Veri
# ----------------------------------------------------------------------------

@dataclass
class Alim:
    satir: int
    tarih: date
    tedarikci: str        # birleştirilmiş görünen ad
    ted_anahtar: str
    kategori: str
    malzeme: str
    miktar: Decimal | None
    birim: str
    fiyat_tl: Decimal | None
    tutar_tl: Decimal
    doviz: str
    departman: str


def oku(yol: Path) -> tuple[list[Alim], list[dict]]:
    uy, ham = [], []
    kursuz = defaultdict(lambda: [0, SIFIR])
    for r in kayitlar(yol, SUTUNLAR, ("tedarikci",)):
        ted = metin(r.get("tedarikci"))
        t = tarih(r.get("tarih"))
        if not ted:
            continue
        if not t:
            uy.append({"onem": "Orta", "tur": "Tarih okunamadı", "kim": ted, "aciklama": f"Satır {r['_satir']} alınmadı"})
            continue
        miktar, fiyat, tutar = para(r.get("miktar")), para(r.get("fiyat")), para(r.get("tutar"))
        if tutar is None and miktar is not None and fiyat is not None:
            tutar = miktar * fiyat
        if tutar is None:
            uy.append({"onem": "Orta", "tur": "Tutar yok", "kim": ted, "aciklama": f"Satır {r['_satir']}: tutar veya miktar × fiyat yok; alınmadı"})
            continue
        doviz = (metin(r.get("doviz")) or "TRY").upper().replace("TL", "TRY")
        kur = Decimal(1) if doviz == "TRY" else para(r.get("kur"))
        if kur is None:
            kursuz[doviz][0] += 1
            kursuz[doviz][1] += tutar
            continue
        if fiyat is None and miktar:
            fiyat = tutar / miktar
        ham.append((r["_satir"], t, ted, metin(r.get("kategori")), metin(r.get("malzeme")), miktar, metin(r.get("birim")), fiyat * kur if fiyat is not None else None,
                    tutar * kur, doviz, metin(r.get("departman"))))
    for dv, (n, t) in kursuz.items():
        uy.append({"onem": "Yüksek", "tur": "Kur yok", "kim": dv, "aciklama": f"{n} satır ({tl(t)} {dv}) kuru olmadığı için analize alınmadı"})
    # Tedarikçi adlarını birleştir
    yazilis = defaultdict(Counter)
    for h in ham:
        yazilis[tedarikci_anahtari(h[2])][h[2]] += 1
    gorunen = {k: c.most_common(1)[0][0] for k, c in yazilis.items()}
    for k, c in yazilis.items():
        if len(c) > 1:
            uy.append({"onem": "Bilgi", "tur": "Tedarikçi adı birleştirildi", "kim": gorunen[k], "aciklama": "Yazılışlar: " + " | ".join(sorted(c))})
    alimlar = []
    kategorisiz = 0
    for (sat, t, ted, kat, mal, mik, bir, fiy, tut, dv, dep) in ham:
        if not kat:
            kategorisiz += 1
        k = tedarikci_anahtari(ted)
        alimlar.append(Alim(sat, t, gorunen[k], k, kat or "Sınıflandırılmamış", mal, mik, bir, fiy, tut, dv, dep))
    if kategorisiz:
        uy.append({"onem": "Orta", "tur": "Kategorisiz harcama", "kim": "Sınıflandırılmamış",
                   "aciklama": f"{kategorisiz} satırın kategorisi yok; ABC ve pazarlık analizinde ayrı grup olarak gösterildi"})
    return alimlar, uy


# ----------------------------------------------------------------------------
# Analiz
# ----------------------------------------------------------------------------

def abc(tutarlar: dict[str, Decimal], a_esik: Decimal, b_esik: Decimal) -> list[dict]:
    toplam = sum(tutarlar.values(), SIFIR)
    sonuc, kum = [], SIFIR
    for ad, t in sorted(tutarlar.items(), key=lambda i: (-i[1], katla(i[0]))):
        onceki = kum / toplam if toplam else SIFIR
        kum += t
        sinif = "A" if onceki < a_esik else "B" if onceki < b_esik else "C"
        sonuc.append({"ad": ad, "tutar": t, "pay": t / toplam if toplam else SIFIR, "kum": kum / toplam if toplam else SIFIR, "sinif": sinif})
    return sonuc


def donem_adi(d: date, tur: str) -> str:
    return ceyrek(d) if tur == "ceyrek" else f"{d.year}-{d.month:02d}"


def fiyat_farklari(alimlar: list[Alim], donem: str = "ay") -> list[dict]:
    gruplar = defaultdict(list)
    for a in alimlar:
        if a.miktar and a.miktar > 0 and a.fiyat_tl is not None and a.malzeme:
            gruplar[(katla(a.malzeme), katla(a.birim), donem_adi(a.tarih, donem))].append(a)
    sonuc = []
    for (_, _, c), lst in gruplar.items():
        fiyatlar = {a.fiyat_tl.quantize(K2) for a in lst}
        if len(lst) < 2 or len(fiyatlar) < 2:
            continue
        mik = sum((a.miktar for a in lst), SIFIR)
        tut = sum((a.fiyat_tl * a.miktar for a in lst), SIFIR)
        ort, enk, enb = tut / mik, min(a.fiyat_tl for a in lst), max(a.fiyat_tl for a in lst)
        sonuc.append({"malzeme": lst[0].malzeme, "birim": lst[0].birim, "donem": c, "kategori": Counter(a.kategori for a in lst).most_common(1)[0][0],
                      "alim": len(lst), "miktar": mik, "tutar": tut, "ort": ort, "min": enk, "max": enb, "fark": (enb - enk) / enk if enk else None,
                      "tedarikciler": sorted({a.tedarikci for a in lst}), "en_ucuz": min(lst, key=lambda a: (a.fiyat_tl, a.tarih)).tedarikci,
                      "tasarruf_ort": sum((max(SIFIR, a.fiyat_tl - ort) * a.miktar for a in lst), SIFIR),
                      "tasarruf_min": sum(((a.fiyat_tl - enk) * a.miktar for a in lst), SIFIR)})
    return sorted(sonuc, key=lambda x: -x["tasarruf_ort"])


def fiyat_artislari(alimlar: list[Alim], esik: Decimal) -> list[dict]:
    """İlk ve son ay ağırlıklı ortalama fiyat. Enflasyon ortamında mutlak artış yanıltıcı olduğundan, artışı tüm
    malzemelerin medyan artışıyla karşılaştırır: (değişim − medyan) > esik ise işaretlenir."""
    gruplar = defaultdict(lambda: defaultdict(lambda: [SIFIR, SIFIR]))
    adlar = {}
    for a in alimlar:
        if a.miktar and a.miktar > 0 and a.fiyat_tl is not None and a.malzeme:
            k = (katla(a.malzeme), katla(a.birim))
            adlar[k] = (a.malzeme, a.birim, a.kategori)
            g = gruplar[k][(a.tarih.year, a.tarih.month)]
            g[0] += a.fiyat_tl * a.miktar
            g[1] += a.miktar
    sonuc = []
    for k, aylar in gruplar.items():
        if len(aylar) < 2:
            continue
        ilk, son = min(aylar), max(aylar)
        f1, f2 = aylar[ilk][0] / aylar[ilk][1], aylar[son][0] / aylar[son][1]
        degisim = f2 / f1 - 1 if f1 else None
        sonuc.append({"malzeme": adlar[k][0], "birim": adlar[k][1], "kategori": adlar[k][2], "ilk_ay": f"{AYLAR[ilk[1] - 1]} {ilk[0]}", "ilk": f1,
                      "son_ay": f"{AYLAR[son[1] - 1]} {son[0]}", "son": f2, "degisim": degisim})
    d = sorted(x["degisim"] for x in sonuc if x["degisim"] is not None)
    medyan = (d[len(d) // 2] if len(d) % 2 else (d[len(d) // 2 - 1] + d[len(d) // 2]) / 2) if d else None
    for x in sonuc:
        x["medyan"] = medyan
        x["fark"] = x["degisim"] - medyan if x["degisim"] is not None and medyan is not None else None
        x["esik_ustu"] = len(d) >= 3 and x["fark"] is not None and x["fark"] > esik
    return sorted(sonuc, key=lambda x: -(x["degisim"] or 0))


def analiz_et(alimlar: list[Alim], a_esik: Decimal = Decimal("0.80"), b_esik: Decimal = Decimal("0.95"), artis: Decimal = Decimal("0.10"),
              tek_kaynak: Decimal = Decimal("0.80"), dagitik: int = 5, fiyat_donem: str = "ay") -> dict:
    toplam = sum((a.tutar_tl for a in alimlar), SIFIR)
    kat_t, ted_t = defaultdict(Decimal), defaultdict(Decimal)
    kat_ted = defaultdict(lambda: defaultdict(Decimal))
    kat_ay = defaultdict(lambda: defaultdict(Decimal))
    for a in alimlar:
        kat_t[a.kategori] += a.tutar_tl
        ted_t[a.tedarikci] += a.tutar_tl
        kat_ted[a.kategori][a.tedarikci] += a.tutar_tl
        kat_ay[a.kategori][(a.tarih.year, a.tarih.month)] += a.tutar_tl
    kat_abc = abc(kat_t, a_esik, b_esik)
    ted_abc = abc(ted_t, a_esik, b_esik)
    ted_sinif = {x["ad"]: x["sinif"] for x in ted_abc}
    kat_sinif = {x["ad"]: x["sinif"] for x in kat_abc}
    ff = fiyat_farklari(alimlar, fiyat_donem)
    fa = fiyat_artislari(alimlar, artis)

    kt = []
    for kat, d in kat_ted.items():
        kt_top = sum(d.values(), SIFIR)
        en_buyuk = max(d.items(), key=lambda i: (i[1], katla(i[0])))
        kucuk = [t for t in d if ted_sinif[t] == "C"]
        kt.append({"kategori": kat, "sinif": kat_sinif[kat], "tutar": kt_top, "tedarikci_sayisi": len(d), "en_buyuk": en_buyuk[0], "en_buyuk_pay": en_buyuk[1] / kt_top,
                   "kucuk": len(kucuk), "kucuk_tutar": sum((d[t] for t in kucuk), SIFIR),
                   "tasarruf_ort": sum((x["tasarruf_ort"] for x in ff if x["kategori"] == kat), SIFIR),
                   "dagilim": sorted(d.items(), key=lambda i: -i[1])})
    kt.sort(key=lambda x: -x["tutar"])

    # Pazarlık öncelikleri
    oneriler = []
    for x in kt:
        if x["kategori"] == "Sınıflandırılmamış":
            continue
        s = x["sinif"]
        if x["tasarruf_ort"] > 0 and x["tedarikci_sayisi"] >= 2:
            oran_yazi = f"{x['tasarruf_ort'] / x['tutar'] * 100:.1f}".replace(".", ",")
            oneriler.append({"oncelik": 1 if s == "A" else 2 if s == "B" else 3, "tur": "Fiyat birliği / toplu pazarlık", "kapsam": x["kategori"],
                             "gerekce": f"Aynı malzeme aynı {'çeyrekte' if fiyat_donem == 'ceyrek' else 'ayda'} farklı fiyatlarla alınmış; ortalamanın üzerindeki alımların fazlası {tl(x['tasarruf_ort'])} TL "
                                        f"(kategori harcamasına oranı %{oran_yazi})", "potansiyel": x["tasarruf_ort"], "tutar": x["tutar"]})
        if x["en_buyuk_pay"] >= tek_kaynak and s in ("A", "B"):
            oneriler.append({"oncelik": 1 if s == "A" else 2, "tur": "Tek kaynak bağımlılığı", "kapsam": x["kategori"],
                             "gerekce": f"{x['en_buyuk']} kategori payı %{x['en_buyuk_pay'] * 100:.0f}; çerçeve sözleşme / hacim indirimi "
                                        "pazarlığı ve alternatif tedarikçi değerlendirmesi", "potansiyel": None, "tutar": x["tutar"]})
        if x["tedarikci_sayisi"] >= dagitik and x["kucuk"] >= 3:
            oneriler.append({"oncelik": 2 if s == "A" else 3, "tur": "Tedarikçi konsolidasyonu", "kapsam": x["kategori"],
                             "gerekce": f"{x['tedarikci_sayisi']} tedarikçi; bunlardan C sınıfı (kuyruk) {x['kucuk']} tedarikçiden {tl(x['kucuk_tutar'])} TL alım. "
                                        "Az sayıda tedarikçide toplayarak hacim indirimi isteyin", "potansiyel": None, "tutar": x["tutar"]})
    for y in fa:
        if y["esik_ustu"] and kat_sinif.get(y["kategori"]) in ("A", "B"):
            oneriler.append({"oncelik": 2, "tur": "Fiyat artışı incelemesi", "kapsam": f"{y['malzeme']} ({y['kategori']})",
                             "gerekce": f"Birim fiyat {y['ilk_ay']} {tl(y['ilk'])} TL → {y['son_ay']} {tl(y['son'])} TL (%{y['degisim'] * 100:+.0f}; tüm malzemelerin "
                                        f"medyan artışı %{y['medyan'] * 100:+.0f}). Endeks / hammadde fiyatıyla karşılaştırın", "potansiyel": None, "tutar": None})
    oneriler.sort(key=lambda o: (o["oncelik"], -(o["potansiyel"] or 0), -(o["tutar"] or 0)))

    uy = []
    for y in fa:
        if y["esik_ustu"]:
            uy.append({"onem": "Orta", "tur": "Fiyat artışı", "kim": y["malzeme"],
                       "aciklama": f"{y['ilk_ay']} → {y['son_ay']}: {tl(y['ilk'])} → {tl(y['son'])} TL (%{y['degisim'] * 100:+.0f}); medyan artış "
                                   f"%{y['medyan'] * 100:+.0f}, fark {y['fark'] * 100:+.0f} puan (eşik {artis * 100:.0f} puan)"})
    for x in kt:
        if x["en_buyuk_pay"] >= tek_kaynak and x["sinif"] == "A":
            uy.append({"onem": "Orta", "tur": "Tek kaynak", "kim": x["kategori"], "aciklama": f"{x['en_buyuk']} payı %{x['en_buyuk_pay'] * 100:.0f}"})
    aylar = sorted({(a.tarih.year, a.tarih.month) for a in alimlar})
    return {"alimlar": alimlar, "toplam": toplam, "kat_abc": kat_abc, "ted_abc": ted_abc, "ted_sinif": ted_sinif, "kat_ted": kt, "ff": ff, "fa": fa,
            "oneriler": oneriler, "uyarilar": uy, "kat_ay": kat_ay, "aylar": aylar, "esik": (a_esik, b_esik), "artis": artis}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
RENK = {"Yüksek": "FDE2E1", "Orta": "FFF4CE", "Bilgi": "E8F0FE"}
SINIF_RENK = {"A": "FDE2E1", "B": "FFF4CE", "C": "E8F0FE"}
ONCELIK_RENK = {1: "FDE2E1", 2: "FFF4CE", 3: "E8F0FE"}
UST = Alignment(vertical="top", wrap_text=True)
PF, YF = "#,##0", "0.0%"


def _baslik(ws, basliklar, genislik=()):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "B2"


def _f(x):
    return None if x is None else float(x)


def _abc_sayfa(wb, ad, satirlar, etiket):
    ws = wb.create_sheet(ad)
    _baslik(ws, [etiket, "Harcama (TL)", "Pay", "Kümülatif Pay", "Sınıf"], (34, 15, 9, 12, 7))
    for x in satirlar:
        ws.append([x["ad"], float(x["tutar"]), float(x["pay"]), float(x["kum"]), x["sinif"]])
        ws.cell(ws.max_row, 2).number_format = PF
        ws.cell(ws.max_row, 3).number_format = ws.cell(ws.max_row, 4).number_format = YF
        ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor=SINIF_RENK[x["sinif"]])
    ozet = defaultdict(lambda: [0, SIFIR])
    for x in satirlar:
        ozet[x["sinif"]][0] += 1
        ozet[x["sinif"]][1] += x["tutar"]
    toplam = sum((x["tutar"] for x in satirlar), SIFIR)
    ws.append([])
    ws.append(["Sınıf", "Sayı", "Sayı Payı", "Harcama", "Harcama Payı"])
    for h in ws[ws.max_row]:
        h.font = Font(bold=True)
    for s_ in "ABC":
        n, t = ozet[s_]
        ws.append([s_, n, n / len(satirlar) if satirlar else 0, float(t), float(t / toplam) if toplam else 0])
        ws.cell(ws.max_row, 3).number_format = ws.cell(ws.max_row, 5).number_format = YF
        ws.cell(ws.max_row, 4).number_format = PF
    return ws


def rapor_yaz(cikti: Path, s: dict) -> None:
    wb = Workbook()
    oz = wb.active
    oz.title = "Özet"
    _baslik(oz, ["Gösterge", "Değer"], (46, 20))
    a = s["alimlar"]
    for k, v in [("Dönem", f"{min(x.tarih for x in a):%d.%m.%Y} – {max(x.tarih for x in a):%d.%m.%Y}"), ("Toplam harcama (TL, KDV hariç)", float(s["toplam"])),
                 ("Satır sayısı", len(a)), ("Kategori sayısı", len(s["kat_abc"])), ("Tedarikçi sayısı", len(s["ted_abc"])),
                 ("A sınıfı tedarikçi sayısı", sum(x["sinif"] == "A" for x in s["ted_abc"])),
                 ("A sınıfı tedarikçilerin harcama payı", float(sum((x["pay"] for x in s["ted_abc"] if x["sinif"] == "A"), SIFIR))),
                 ("C sınıfı (kuyruk) tedarikçi sayısı", sum(x["sinif"] == "C" for x in s["ted_abc"])),
                 ("Fiyat birliği potansiyeli (ortalamanın üstü, TL)", float(sum((x["tasarruf_ort"] for x in s["ff"]), SIFIR))),
                 ("Fiyat birliği potansiyeli (en düşük fiyata göre, üst sınır, TL)", float(sum((x["tasarruf_min"] for x in s["ff"]), SIFIR))),
                 ("1. öncelikli pazarlık konusu", sum(o["oncelik"] == 1 for o in s["oneriler"])),
                 ("ABC sınırları", f"A < %{s['esik'][0] * 100:g}, B < %{s['esik'][1] * 100:g} (önceki kümülatif pay)")]:
        oz.append([k, v])
        if isinstance(v, float):
            oz.cell(oz.max_row, 2).number_format = YF if "payı" in k else PF

    po = wb.create_sheet("Pazarlık Öncelikleri")
    _baslik(po, ["Öncelik", "Tür", "Kapsam", "Gerekçe", "Potansiyel (TL)", "Kategori Harcaması", "Sorumlu", "Hedef Tarih", "Sonuç"],
            (8, 26, 30, 80, 13, 15, 14, 11, 20))
    for o in s["oneriler"]:
        po.append([o["oncelik"], o["tur"], o["kapsam"], o["gerekce"], _f(o["potansiyel"]), _f(o["tutar"]), "", "", ""])
        po.cell(po.max_row, 1).fill = PatternFill("solid", fgColor=ONCELIK_RENK[o["oncelik"]])
        po.cell(po.max_row, 4).alignment = UST
        po.cell(po.max_row, 5).number_format = po.cell(po.max_row, 6).number_format = PF
        for j in (7, 8, 9):
            po.cell(po.max_row, j).fill = PatternFill("solid", fgColor="FFF4CE")
    ka = _abc_sayfa(wb, "Kategori ABC", s["kat_abc"], "Kategori")
    g = BarChart()
    g.title, g.height, g.width = "Kategori harcaması", 8, 18
    g.add_data(Reference(ka, min_col=2, min_row=1, max_row=len(s["kat_abc"]) + 1), titles_from_data=True)
    g.set_categories(Reference(ka, min_col=1, min_row=2, max_row=len(s["kat_abc"]) + 1))
    ka.add_chart(g, "H2")
    _abc_sayfa(wb, "Tedarikçi ABC", s["ted_abc"], "Tedarikçi")

    kt = wb.create_sheet("Kategori × Tedarikçi")
    _baslik(kt, ["Kategori", "Sınıf", "Harcama", "Tedarikçi Sayısı", "En Büyük Tedarikçi", "En Büyük Pay", "C Sınıfı Tedarikçi", "Bunların Tutarı",
                 "Tedarikçi Dağılımı"], (24, 6, 14, 10, 26, 10, 11, 13, 90))
    for x in s["kat_ted"]:
        kt.append([x["kategori"], x["sinif"], float(x["tutar"]), x["tedarikci_sayisi"], x["en_buyuk"], float(x["en_buyuk_pay"]), x["kucuk"], float(x["kucuk_tutar"]),
                   "; ".join(f"{t} %{v / x['tutar'] * 100:.0f}" for t, v in x["dagilim"])])
        kt.cell(kt.max_row, 3).number_format = kt.cell(kt.max_row, 8).number_format = PF
        kt.cell(kt.max_row, 6).number_format = "0%"

    ff = wb.create_sheet("Fiyat Farkları")
    _baslik(ff, ["Malzeme", "Birim", "Dönem", "Kategori", "Alım", "Miktar", "Ort. Fiyat", "En Düşük", "En Yüksek", "Fark", "En Ucuz Tedarikçi", "Tedarikçiler",
                 "Ortalamanın Üstü Fazla (TL)", "En Düşüğe Göre Fark (TL)"], (28, 7, 9, 18, 6, 9, 10, 10, 10, 7, 22, 40, 14, 14))
    for x in s["ff"]:
        ff.append([x["malzeme"], x["birim"], x["donem"], x["kategori"], x["alim"], float(x["miktar"]), float(x["ort"]), float(x["min"]), float(x["max"]), _f(x["fark"]),
                   x["en_ucuz"], ", ".join(x["tedarikciler"]), float(x["tasarruf_ort"]), float(x["tasarruf_min"])])
        for j in (7, 8, 9):
            ff.cell(ff.max_row, j).number_format = "#,##0.00"
        ff.cell(ff.max_row, 10).number_format = YF
        ff.cell(ff.max_row, 13).number_format = ff.cell(ff.max_row, 14).number_format = PF

    fa = wb.create_sheet("Fiyat Artışları")
    _baslik(fa, ["Malzeme", "Birim", "Kategori", "İlk Ay", "İlk Ay Fiyatı", "Son Ay", "Son Ay Fiyatı", "Değişim", "Medyan Değişim", "Fark (puan)"],
            (28, 7, 18, 13, 12, 13, 12, 9, 10, 9))
    for y in s["fa"]:
        fa.append([y["malzeme"], y["birim"], y["kategori"], y["ilk_ay"], float(y["ilk"]), y["son_ay"], float(y["son"]), _f(y["degisim"]), _f(y["medyan"]),
                   None if y["fark"] is None else round(float(y["fark"]) * 100, 1)])
        fa.cell(fa.max_row, 5).number_format = fa.cell(fa.max_row, 7).number_format = "#,##0.00"
        fa.cell(fa.max_row, 8).number_format = fa.cell(fa.max_row, 9).number_format = YF
        if y["esik_ustu"]:
            fa.cell(fa.max_row, 10).fill = PatternFill("solid", fgColor="FFF4CE")

    at = wb.create_sheet("Aylık Trend")
    aylar = s["aylar"]
    _baslik(at, ["Kategori"] + [f"{AYLAR[m - 1][:3]} {y}" for y, m in aylar] + ["Toplam"], [24] + [11] * (len(aylar) + 1))
    for x in s["kat_abc"]:
        d = s["kat_ay"][x["ad"]]
        at.append([x["ad"]] + [float(d.get(ay, SIFIR)) for ay in aylar] + [float(x["tutar"])])
        for j in range(2, len(aylar) + 3):
            at.cell(at.max_row, j).number_format = PF

    vr = wb.create_sheet("Veri")
    _baslik(vr, ["Tarih", "Tedarikçi", "Kategori", "Malzeme", "Miktar", "Birim", "Birim Fiyat (TL)", "Tutar (TL)", "Döviz", "Departman", "Tedarikçi Sınıfı"],
            (11, 26, 20, 28, 9, 7, 12, 13, 6, 14, 9))
    for x in sorted(a, key=lambda x: (x.tarih, x.satir)):
        vr.append([x.tarih, x.tedarikci, x.kategori, x.malzeme, _f(x.miktar), x.birim, _f(x.fiyat_tl), float(x.tutar_tl), x.doviz, x.departman, s["ted_sinif"][x.tedarikci]])
        vr.cell(vr.max_row, 1).number_format = "DD.MM.YYYY"
        vr.cell(vr.max_row, 7).number_format = "#,##0.00"
        vr.cell(vr.max_row, 8).number_format = PF
    vr.auto_filter.ref = f"A1:K{vr.max_row}"

    uy = wb.create_sheet("Uyarılar")
    _baslik(uy, ["Önem", "Tür", "Kalem / Kim", "Açıklama"], (9, 26, 30, 100))
    for u in s["uyarilar"]:
        uy.append([u["onem"], u["tur"], u["kim"], u["aciklama"]])
        uy.cell(uy.max_row, 1).fill = PatternFill("solid", fgColor=RENK[u["onem"]])
        uy.cell(uy.max_row, 4).alignment = UST
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(yol: Path, cikti: Path, **kw) -> dict:
    alimlar, uy = oku(yol)
    if not alimlar:
        raise ValueError(f"{yol.name}: satın alma satırı bulunamadı")
    s = analiz_et(alimlar, **kw)
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    s["uyarilar"] = sorted(uy + s["uyarilar"], key=lambda u: (sira[u["onem"]], u["tur"], katla(u["kim"])))
    rapor_yaz(cikti, s)
    return s


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Satın alma harcamasını kategori ve tedarikçi bazında ABC ile analiz eder, pazarlık önceliklerini çıkarır.")
    p.add_argument("--veri", type=Path, default=ORNEK / "satinalma.csv",
                   help="Tarih, Tedarikçi, Kategori, Malzeme, Miktar, Birim, Birim Fiyat, Tutar (KDV hariç), Döviz, Kur, Departman")
    p.add_argument("--a-esik", type=float, default=80, help="A sınıfı kümülatif pay sınırı %% (varsayılan 80)")
    p.add_argument("--b-esik", type=float, default=95, help="B sınıfı kümülatif pay sınırı %% (varsayılan 95)")
    p.add_argument("--artis", type=float, default=10, help="Birim fiyat artışı tüm malzemelerin medyan artışını bu kadar puan aşarsa uyarı (varsayılan 10)")
    p.add_argument("--tek-kaynak", type=float, default=80, help="En büyük tedarikçinin kategori payı bu %%'ye ulaşırsa tek kaynak (varsayılan 80)")
    p.add_argument("--dagitik", type=int, default=5, help="Kategoride bu kadar ve daha fazla tedarikçi varsa konsolidasyon değerlendirilir (varsayılan 5)")
    p.add_argument("--fiyat-donem", choices=["ay", "ceyrek"], default="ay", help="Fiyat farkı hangi dönem içinde karşılaştırılsın (varsayılan ay)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "harcama_analizi.xlsx")
    a = p.parse_args(argv)
    if not a.veri.exists():
        print(f"[X] Dosya bulunamadı: {a.veri}")
        return 1
    if not 0 < a.a_esik < a.b_esik < 100:
        print("[X] 0 < --a-esik < --b-esik < 100 olmalı")
        return 2
    try:
        s = calistir(a.veri, a.cikti, a_esik=Decimal(str(a.a_esik)) / 100, b_esik=Decimal(str(a.b_esik)) / 100, artis=Decimal(str(a.artis)) / 100,
                     tek_kaynak=Decimal(str(a.tek_kaynak)) / 100, dagitik=a.dagitik, fiyat_donem=a.fiyat_donem)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    ta = [x for x in s["ted_abc"] if x["sinif"] == "A"]
    print(f"[OK] Toplam {tl(s['toplam'])} TL · {len(s['kat_abc'])} kategori · {len(s['ted_abc'])} tedarikçi; "
          f"A sınıfı {len(ta)} tedarikçinin payı %{sum((x['pay'] for x in ta), SIFIR) * 100:.0f}")
    print(f"[OK] Fiyat birliği potansiyeli {tl(sum((x['tasarruf_ort'] for x in s['ff']), SIFIR))} TL (ortalamanın üstü)")
    for o in s["oneriler"][:6]:
        print(f"     {o['oncelik']}. {o['tur']} — {o['kapsam']}")
    for u in s["uyarilar"]:
        if u["onem"] == "Yüksek":
            print(f"[!] {u['kim']} · {u['tur']}: {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
