"""
Araç ve Sefer Takip Listesi — Workers / Workless kod bloğu
Lojistik ve Taşımacılık › Karayolu Operasyon › Karayolu Operasyon Uzmanı

Uluslararası karayolu seferlerini yükleme, çıkış gümrüğü, sınır çıkışı, varış gümrüğü ve teslim aşamalarıyla
takip eder:
  - Tahmini teslim: son gerçekleşen aşamadan itibaren varış ülkesine göre standart aşama süreleri (saat). Süresi
    geçmiş ama gerçekleşmemiş aşamanın en erken şimdi gerçekleşeceği varsayılır (ör. sınır kuyruğu).
  - Yüklemesi gecikmiş, aşaması gecikmiş, planlanan teslimi kaçıracak seferler (24 saatten fazla gecikme Yüksek).
  - Konum bilgisi eski (varsayılan 12 saat), yüklenmiş seferde CMR numarası yok.
  - Araç çakışması: aynı çekicinin yeni seferi, önceki seferin tahmini tesliminden önce planlanmış.
Rapor: sefer listesi, araç durumu (müsait olacağı zaman), aşama matrisi, müşteri / ülke performansı, uyarılar.
İnternete bağlanmaz.

Kullanım:
    python main.py                                                       # örnek: 8 sefer, şimdi 08.10.2026 09:00
    python main.py --seferler seferler.xlsx --sureler asama_sureleri.csv --simdi "08.10.2026 09:00"
"""
from __future__ import annotations

import argparse
import csv
import re
import statistics
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
ASAMALAR = ["Yükleme", "Çıkış Gümrüğü", "Sınır Çıkış", "Varış Gümrüğü", "Teslim"]
ASAMA_ADLARI = {"Yükleme": ("yukleme", "yukleme tarihi"), "Çıkış Gümrüğü": ("cikis gumrugu", "ihracat gumrugu", "gumruk cikis"),
                "Sınır Çıkış": ("sinir cikis", "sinir", "sinir kapisi cikis"), "Varış Gümrüğü": ("varis gumrugu", "ithalat gumrugu", "gumruk varis"),
                "Teslim": ("teslim", "teslim tarihi", "bosaltma")}
SEFER_SUTUNLARI = {"no": ("sefer no", "sefer", "is no"), "cekici": ("cekici plaka", "plaka", "cekici"), "dorse": ("dorse plaka", "dorse"),
                   "sofor": ("sofor", "surucu"), "musteri": ("musteri", "gonderici"), "ulke": ("varis ulke", "ulke", "varis ulkesi"),
                   "sehir": ("varis sehri", "sehir", "varis"), "kapi": ("sinir kapisi", "kapi", "gumruk kapisi"),
                   "p_yukleme": ("planlanan yukleme", "plan yukleme", "yukleme plani"), "p_teslim": ("planlanan teslim", "plan teslim", "teslim termini"),
                   "cmr": ("cmr no", "cmr"), "konum": ("son konum", "konum"), "konum_zaman": ("son konum zamani", "konum zamani", "konum tarihi")}
SURE_SUTUNLARI = {"ulke": ("varis ulke", "ulke", "guzergah"), "asama": ("asama",), "sure": ("sure saat", "sure", "saat")}


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(s or "").translate(_TR).lower()).split())


def zaman(x) -> datetime | None:
    if isinstance(x, datetime):
        return x
    if isinstance(x, date):
        return datetime.combine(x, time(0, 0))
    s = str(x or "").strip()
    for f in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H.%M", "%Y-%m-%d %H:%M", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    return None


def sayi(x) -> float | None:
    if x in (None, ""):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    try:
        return float(str(x).strip().replace(",", "."))
    except ValueError:
        return None


def metin(x) -> str:
    return str(x if x is not None else "").strip()


def saat(td: timedelta) -> float:
    return td.total_seconds() / 3600


def sure_yaz(saatler: float) -> str:
    s = round(abs(saatler))
    return f"{s // 24} gün {s % 24} saat" if s >= 24 else f"{s} saat"


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
# Veri
# ----------------------------------------------------------------------------

@dataclass
class Sefer:
    satir: int
    no: str
    cekici: str
    dorse: str
    sofor: str
    musteri: str
    ulke: str
    sehir: str
    kapi: str
    p_yukleme: datetime | None
    p_teslim: datetime | None
    cmr: str
    konum: str
    konum_zaman: datetime | None
    gercek: dict                                   # aşama → zaman
    tahmini: dict = field(default_factory=dict)
    mevcut: str = ""
    tahmini_teslim: datetime | None = None
    fark_saat: float | None = None                 # + gecikme
    durum: str = ""

    @property
    def teslim(self) -> datetime | None:
        return self.gercek.get("Teslim")

    @property
    def bitis(self) -> datetime | None:
        return self.teslim or self.tahmini_teslim

    @property
    def aktif(self) -> bool:
        return bool(self.gercek) and not self.teslim


def sureleri_oku(yol: Path) -> dict[str, dict[str, float]]:
    sonuc = defaultdict(dict)
    for r in kayitlar(yol, SURE_SUTUNLARI, ("ulke", "asama", "sure")):
        asama = next((a for a, adlar in ASAMA_ADLARI.items() if katla(r.get("asama")) in adlar), None)
        if asama and sayi(r.get("sure")) is not None:
            sonuc[katla(r["ulke"])][asama] = sayi(r["sure"])
    return dict(sonuc)


def seferleri_oku(yol: Path) -> list[Sefer]:
    s = tablo_oku(yol)
    basliklar = next((r for r in s[:15] if "sefer no" in [katla(c) for c in r] or "sefer" in [katla(c) for c in r]), None)
    asama_j = {}
    if basliklar:
        b = [katla(c) for c in basliklar]
        for a, adlar in ASAMA_ADLARI.items():
            j = next((b.index(x) for x in adlar if x in b), None)
            if j is not None:
                asama_j[a] = j
    sonuc = []
    for r in kayitlar(yol, SEFER_SUTUNLARI, ("no",)):
        if not r.get("no"):
            continue
        ham = s[r["_satir"] - 1] if r["_satir"] - 1 < len(s) else []
        gercek = {a: zaman(ham[j]) for a, j in asama_j.items() if j < len(ham) and zaman(ham[j])}
        sonuc.append(Sefer(r["_satir"], metin(r["no"]), metin(r.get("cekici")), metin(r.get("dorse")), metin(r.get("sofor")), metin(r.get("musteri")),
                           metin(r.get("ulke")), metin(r.get("sehir")), metin(r.get("kapi")), zaman(r.get("p_yukleme")), zaman(r.get("p_teslim")),
                           metin(r.get("cmr")), metin(r.get("konum")), zaman(r.get("konum_zaman")), gercek))
    return sonuc


# ----------------------------------------------------------------------------
# Analiz
# ----------------------------------------------------------------------------

SIRA = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}


def analiz_et(seferler: list[Sefer], sureler: dict, simdi: datetime, konum_saat: float = 12, yuksek_saat: float = 24) -> dict:
    uyarilar = []

    def uyar(onem, tur, s: Sefer, aciklama):
        uyarilar.append({"onem": onem, "tur": tur, "sefer": s.no, "cekici": s.cekici, "musteri": s.musteri, "aciklama": aciklama})

    for s in seferler:
        sablon = sureler.get(katla(s.ulke)) or sureler.get("")          # "*" satırı katlanınca ""
        # Aşama sırası
        onceki = None
        for a in ASAMALAR:
            t = s.gercek.get(a)
            if t and onceki and t < onceki[1]:
                uyar("Orta", "Veri hatası", s, f"{a} ({t:%d.%m.%Y %H:%M}) {onceki[0]} aşamasından ({onceki[1]:%d.%m.%Y %H:%M}) önce")
            if t:
                onceki = (a, t)
        if s.teslim:
            s.durum, s.mevcut = "Teslim edildi", "Teslim edildi"
            if s.p_teslim:
                s.fark_saat = saat(s.teslim - s.p_teslim)
                if s.fark_saat > 0:
                    uyar("Bilgi", "Geç teslim", s, f"{s.teslim:%d.%m.%Y %H:%M} teslim edildi, plan {s.p_teslim:%d.%m.%Y %H:%M} ({sure_yaz(s.fark_saat)} geç)")
            continue
        if sablon is None:
            s.durum = "Süre tanımı yok"
            uyar("Orta", "Süre tanımı yok", s, f"'{s.ulke}' için aşama süresi tanımlı değil; tahmini teslim hesaplanamadı")
            continue
        sira = [a for a in ASAMALAR if a == "Yükleme" or a in sablon]
        son_i = max((i for i, a in enumerate(sira) if a in s.gercek), default=-1)
        if son_i < 0:
            if not s.p_yukleme:
                s.durum = "Plan yok"
                uyar("Bilgi", "Yükleme planı yok", s, "Planlanan yükleme zamanı boş")
                continue
            imlec = s.p_yukleme
            if imlec < simdi:
                gec = saat(simdi - imlec)
                uyar("Yüksek" if gec > yuksek_saat else "Orta", "Yükleme gecikti", s, f"Planlanan yükleme {imlec:%d.%m.%Y %H:%M}, {sure_yaz(gec)} geçti; "
                     "yükleme kaydı yok. Yükleme yeriyle / şoförle teyit edin.")
                imlec = simdi
            s.tahmini["Yükleme"] = imlec
            s.mevcut = "Yükleme bekleniyor"
            son_i = 0
        else:
            imlec = s.gercek[sira[son_i]]
        ilk = not s.mevcut
        for a in sira[son_i + 1:]:
            bitis = imlec + timedelta(hours=sablon[a])
            if ilk:
                s.mevcut = f"{sira[sira.index(a) - 1]} tamam → {a} bekleniyor"
                if bitis < simdi:
                    gec = saat(simdi - bitis)
                    ek = f" ({s.kapi})" if a == "Sınır Çıkış" and s.kapi else ""
                    uyar("Orta", "Aşama gecikti", s, f"{a}{ek}: {imlec:%d.%m.%Y %H:%M} + {sablon[a]:g} saat = {bitis:%d.%m.%Y %H:%M} bekleniyordu, "
                         f"{sure_yaz(gec)} geçti" + (f"; son konum: {s.konum}" if s.konum else ""))
                    bitis = simdi
                ilk = False
            s.tahmini[a] = bitis
            imlec = bitis
        s.tahmini_teslim = imlec
        if s.gercek:
            if s.konum_zaman is None or saat(simdi - s.konum_zaman) > konum_saat:
                uyar("Orta", "Konum bilgisi eski", s, "Konum bilgisi yok" if s.konum_zaman is None else
                     f"Son konum '{s.konum}' {s.konum_zaman:%d.%m.%Y %H:%M}, {sure_yaz(saat(simdi - s.konum_zaman))} önce; şoförden güncel konum alın")
            if not s.cmr:
                uyar("Bilgi", "CMR yok", s, "Yükleme yapılmış ama CMR numarası girilmemiş")
        if s.p_teslim:
            s.fark_saat = saat(s.tahmini_teslim - s.p_teslim)
            if s.fark_saat > 0:
                s.durum = "Gecikecek"
                uyar("Yüksek" if s.fark_saat > yuksek_saat else "Orta", "Teslim gecikecek", s,
                     f"Tahmini teslim {s.tahmini_teslim:%d.%m.%Y %H:%M}, plan {s.p_teslim:%d.%m.%Y %H:%M} ({sure_yaz(s.fark_saat)} geç). "
                     "Müşteriyi bilgilendirin.")
            else:
                s.durum = "Yolda" if s.gercek else "Planlandı"
        else:
            s.durum = "Yolda" if s.gercek else "Planlandı"

    # Araç çakışması ve araç durumu
    araclar = defaultdict(list)
    for s in seferler:
        if s.cekici:
            araclar[s.cekici].append(s)
    arac_durumu = []
    for plaka, lst in sorted(araclar.items()):
        lst.sort(key=lambda s: (s.gercek.get("Yükleme") or s.p_yukleme or datetime.max))
        for a, b in zip(lst, lst[1:]):
            bas_b = b.gercek.get("Yükleme") or b.p_yukleme
            if a.bitis and bas_b and bas_b < a.bitis and not b.gercek:
                uyar("Yüksek", "Araç çakışması", b, f"{plaka}: {b.no} yüklemesi {bas_b:%d.%m.%Y %H:%M} planlandı ama aynı araç {a.no} seferinde, tahmini "
                     f"teslim {a.bitis:%d.%m.%Y %H:%M} (varış: {a.sehir}, {a.ulke}); boşaltma sonrası dönüş süresi de gerekir. Başka araç atayın.")
        aktif = next((s for s in lst if s.aktif), None)
        son_teslim = max((s.teslim for s in lst if s.teslim), default=None)
        sirada = [s for s in lst if not s.gercek]
        arac_durumu.append({"plaka": plaka, "aktif": aktif, "sofor": (aktif or lst[-1]).sofor, "konum": aktif.konum if aktif else "",
                            "konum_zaman": aktif.konum_zaman if aktif else None, "musait": aktif.bitis if aktif else son_teslim,
                            "sirada": sirada})

    # Performans
    perf = {"musteri": [], "ulke": []}
    for anahtar, f in (("musteri", lambda s: s.musteri), ("ulke", lambda s: s.ulke)):
        g = defaultdict(list)
        for s in seferler:
            g[f(s) or "—"].append(s)
        for ad, lst in sorted(g.items()):
            biten = [s for s in lst if s.teslim and s.fark_saat is not None]
            gec = [s.fark_saat for s in biten if s.fark_saat > 0]
            perf[anahtar].append({"ad": ad, "sefer": len(lst), "teslim": len(biten), "zamaninda": len(biten) - len(gec),
                                  "acik_gecikecek": sum(s.durum == "Gecikecek" for s in lst), "ort_gec": statistics.fmean(gec) if gec else 0})
    uyarilar.sort(key=lambda u: (SIRA[u["onem"]], u["sefer"]))
    return {"uyarilar": uyarilar, "araclar": arac_durumu, "performans": perf}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
RENK = {"Yüksek": "FDE2E1", "Orta": "FFF4CE", "Bilgi": "E8F0FE", "Gecikecek": "FDE2E1", "Yolda": "E8F0FE", "Teslim edildi": "E3F4E1",
        "Planlandı": "F2F2F2"}
GERCEK = PatternFill("solid", fgColor="E3F4E1")
KONTROL = PatternFill("solid", fgColor="FFF4CE")
UST = Alignment(vertical="top", wrap_text=True)
ZF = "DD.MM.YYYY HH:MM"


def _baslik(ws, basliklar, genislik=()):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def rapor_yaz(cikti: Path, seferler: list[Sefer], s: dict, simdi: datetime) -> None:
    wb = Workbook()
    sl = wb.active
    sl.title = "Seferler"
    _baslik(sl, ["Sefer No", "Çekici", "Dorse", "Şoför", "Müşteri", "Varış", "Sınır Kapısı", "Planlanan Yükleme", "Planlanan Teslim", "Mevcut Durum",
                 "Tahmini / Gerçek Teslim", "Fark (saat)", "Son Konum", "Konum Zamanı", "Durum", "Operasyon Notu"],
            (8, 11, 10, 9, 15, 16, 10, 15, 15, 32, 15, 8, 22, 15, 12, 28))
    sirali = sorted(seferler, key=lambda x: ({"Gecikecek": 0, "Yolda": 1, "Planlandı": 2}.get(x.durum, 3), x.bitis or datetime.max))
    for x in sirali:
        sl.append([x.no, x.cekici, x.dorse, x.sofor, x.musteri, f"{x.sehir}, {x.ulke}".strip(", "), x.kapi, x.p_yukleme, x.p_teslim, x.mevcut, x.bitis,
                   round(x.fark_saat, 1) if x.fark_saat is not None else None, x.konum, x.konum_zaman, x.durum, ""])
        for j in (8, 9, 11, 14):
            sl.cell(sl.max_row, j).number_format = ZF
        if x.durum in RENK:
            sl.cell(sl.max_row, 15).fill = PatternFill("solid", fgColor=RENK[x.durum])
        sl.cell(sl.max_row, 16).fill = KONTROL
    sl.append([])
    sl.append([f"Durum zamanı {simdi:%d.%m.%Y %H:%M}. Tahminler standart aşama süreleriyle hesaplanır."])

    ad = wb.create_sheet("Araç Durumu")
    _baslik(ad, ["Çekici", "Şoför", "Aktif Sefer", "Varış", "Mevcut Durum", "Son Konum", "Konum Zamanı", "Müsait Olacağı Zaman (teslim)",
                 "Sıradaki Seferler"], (11, 9, 9, 16, 32, 22, 15, 16, 34))
    for a in s["araclar"]:
        x = a["aktif"]
        ad.append([a["plaka"], a["sofor"], x.no if x else "Boşta", f"{x.sehir}, {x.ulke}" if x else "", x.mevcut if x else "", a["konum"], a["konum_zaman"],
                   a["musait"], ", ".join(f"{y.no} ({y.p_yukleme:%d.%m %H:%M})" if y.p_yukleme else y.no for y in a["sirada"])])
        ad.cell(ad.max_row, 7).number_format = ad.cell(ad.max_row, 8).number_format = ZF

    am = wb.create_sheet("Aşama Matrisi")
    _baslik(am, ["Sefer No", "Varış"] + ASAMALAR, (8, 18) + (16,) * len(ASAMALAR))
    for x in seferler:
        am.append([x.no, f"{x.sehir}, {x.ulke}".strip(", ")])
        for j, a in enumerate(ASAMALAR, 3):
            c = am.cell(am.max_row, j)
            if a in x.gercek:
                c.value, c.fill = x.gercek[a], GERCEK
            elif a in x.tahmini:
                c.value, c.font = x.tahmini[a], Font(italic=True, color="7F7F7F")
            c.number_format = ZF
    am.append([])
    am.append(["Yeşil: gerçekleşen · gri italik: tahmini"])

    pf = wb.create_sheet("Performans")
    for baslik, anahtar in (("Müşteri", "musteri"), ("Varış Ülke", "ulke")):
        pf.append([baslik, "Sefer", "Teslim Edilen", "Zamanında", "Zamanında %", "Geç Teslimde Ort. (saat)", "Açık ve Gecikecek"])
        for h in pf[pf.max_row]:
            h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        for x in s["performans"][anahtar]:
            pf.append([x["ad"], x["sefer"], x["teslim"], x["zamaninda"], x["zamaninda"] / x["teslim"] if x["teslim"] else None, round(x["ort_gec"], 1),
                       x["acik_gecikecek"]])
            pf.cell(pf.max_row, 5).number_format = "0%"
        pf.append([])
    for j, w in enumerate((18, 7, 10, 10, 10, 12, 12), 1):
        pf.column_dimensions[get_column_letter(j)].width = w

    uy = wb.create_sheet("Uyarılar")
    _baslik(uy, ["Önem", "Tür", "Sefer", "Çekici", "Müşteri", "Açıklama", "Aksiyon"], (9, 20, 8, 11, 15, 90, 24))
    for u in s["uyarilar"]:
        uy.append([u["onem"], u["tur"], u["sefer"], u["cekici"], u["musteri"], u["aciklama"], ""])
        uy.cell(uy.max_row, 1).fill = PatternFill("solid", fgColor=RENK[u["onem"]])
        uy.cell(uy.max_row, 6).alignment = UST
        uy.cell(uy.max_row, 7).fill = KONTROL
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(sefer_yolu: Path, sure_yolu: Path, cikti: Path, simdi: datetime | None = None, konum_saat: float = 12) -> dict:
    simdi = simdi or datetime.now().replace(second=0, microsecond=0)
    sureler = sureleri_oku(sure_yolu)
    seferler = seferleri_oku(sefer_yolu)
    s = analiz_et(seferler, sureler, simdi, konum_saat)
    rapor_yaz(cikti, seferler, s, simdi)
    return {**s, "seferler": seferler, "simdi": simdi}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Karayolu seferlerini yükleme, gümrük, sınır ve teslim aşamalarıyla takip eder; geciken araçları listeler.")
    p.add_argument("--seferler", type=Path, default=ORNEK / "seferler.csv", help="Sefer listesi ve aşama gerçekleşen zamanları (.xlsx/.csv)")
    p.add_argument("--sureler", type=Path, default=ORNEK / "asama_sureleri.csv", help="Varış Ülke, Aşama, Süre (saat)")
    p.add_argument("--simdi", help="Durum zamanı 'GG.AA.YYYY SS:DD' (varsayılan şimdi; örnek veride 08.10.2026 09:00)")
    p.add_argument("--konum-saat", type=float, default=12, help="Konum bilgisi kaç saatten eskiyse uyarılsın (varsayılan 12)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "sefer_takip.xlsx")
    a = p.parse_args(argv)
    for y in (a.seferler, a.sureler):
        if not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    simdi = zaman(a.simdi) if a.simdi else (datetime(2026, 10, 8, 9, 0) if a.seferler == ORNEK / "seferler.csv" else None)
    if a.simdi and simdi is None:
        print("[X] --simdi 'GG.AA.YYYY SS:DD' biçiminde olmalı")
        return 1
    try:
        s = calistir(a.seferler, a.sureler, a.cikti, simdi, a.konum_saat)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    yolda = [x for x in s["seferler"] if x.aktif]
    print(f"[OK] {len(s['seferler'])} sefer · {len(yolda)} yolda · durum {s['simdi']:%d.%m.%Y %H:%M}")
    for u in s["uyarilar"]:
        if u["onem"] == "Yüksek":
            print(f"[!] {u['sefer']} {u['cekici']} {u['tur']}: {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
