"""
Stok Değerleme — Workers / Workless kod bloğu
Muhasebe › Maliyet Muhasebesi Uzmanı

Stok hareketlerinden dönem sonu stok değerini ve satılan malın (kullanılan malzemenin) maliyetini hesaplar.
Üç yöntem yan yana; raporun ana yöntemi --yontem ile seçilir:
  - FIFO (ilk giren ilk çıkar): çıkışlar en eski giriş katmanından karşılanır.
  - Hareketli ağırlıklı ortalama: her girişte ortalama yeniden hesaplanır; çıkış o anki ortalamayla.
  - Dönem sonu (basit) ağırlıklı ortalama: (devir + girişler) ÷ miktar; tüm çıkışlar ve kalan bu birim maliyetle.
Hareket türleri: Devir, Giriş (alış / üretimden giriş), Çıkış, Alış İadesi (fiyatıyla düşülür), Satış İadesi (fiyatsız
giriş: hareketli ortalamada güncel ortalama, FIFO'da son çıkış maliyetiyle geri alınır).
Kontroller: eksi stok (çıkış mevcuttan fazla; eksik kısım son maliyetle maliyetlenir, sonraki giriş önce eksiği kapatır
ve fiyat farkı maliyete düzeltme olarak yansır), fiyatsız giriş, son alış fiyatının ortalamadan %25+ sapması, hareketsiz
stok, isteğe bağlı net gerçekleşebilir değer testi (tahmini satış fiyatı − satış gideri < birim maliyet).
Rapor: değerleme özeti, yöntem karşılaştırması, hareket detayı, FIFO katmanları, değer düşüklüğü, uyarılar.
İnternete bağlanmaz.

Kullanım:
    python main.py                                                 # örnek: 4 stok, Ocak–Mart 2026
    python main.py --hareketler stok_hareketleri.xlsx --yontem hareketli
    python main.py --hareketler h.xlsx --yontem fifo --tarih 31.12.2026 --satis-fiyatlari fiyatlar.xlsx
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
K2 = Decimal("0.01")
SIFIR = Decimal(0)
YONTEMLER = {"fifo": "FIFO", "hareketli": "Hareketli ağırlıklı ortalama", "donemsel": "Dönem sonu ağırlıklı ortalama"}

SUTUNLAR = {"tarih": ("tarih", "islem tarihi", "belge tarihi"), "kod": ("stok kodu", "malzeme kodu", "urun kodu", "kod"),
            "ad": ("stok adi", "malzeme adi", "urun adi", "aciklama"), "birim": ("birim",), "hareket": ("hareket", "hareket turu", "islem turu", "tur"),
            "miktar": ("miktar",), "fiyat": ("birim fiyat", "birim maliyet", "fiyat"), "tutar": ("tutar",), "belge": ("belge no", "fis no", "evrak no")}
FIYAT_SUTUNLARI = {"kod": ("stok kodu", "malzeme kodu", "urun kodu", "kod"), "fiyat": ("tahmini satis fiyati", "satis fiyati", "fiyat"),
                   "gider": ("satis gideri", "tamamlama ve satis gideri", "gider")}


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


def sayi(x) -> Decimal | None:
    if x in (None, ""):
        return None
    if isinstance(x, (int, float)):
        return Decimal(str(x))
    s = str(x).strip().replace("TL", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def yuv(x: Decimal) -> Decimal:
    return x.quantize(K2, ROUND_HALF_UP)


def tl(x: Decimal | None) -> str:
    return "—" if x is None else f"{yuv(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


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


# ----------------------------------------------------------------------------
# Veri
# ----------------------------------------------------------------------------

@dataclass
class Hareket:
    tarih: date
    tur: str                    # devir | giris | cikis | alis_iade | satis_iade
    miktar: Decimal
    fiyat: Decimal | None
    belge: str
    sira: int


def tur_bul(x, fiyat) -> str | None:
    k = katla(x)
    if "alis iade" in k or "satin alma iade" in k:
        return "alis_iade"
    if "satis iade" in k:
        return "satis_iade"
    if k.startswith(("devir", "acilis")):
        return "devir"
    if k.startswith(("giris", "alis", "uretim", "satin")):
        return "giris" if fiyat is not None else "satis_iade"
    if k.startswith(("cikis", "satis", "sarf", "tuketim", "kullanim", "fire")):
        return "cikis"
    return None


@dataclass
class Stok:
    kod: str
    ad: str
    birim: str
    hareketler: list = field(default_factory=list)
    sonuc: dict = field(default_factory=dict)          # yöntem → {miktar, deger, smm, detay, katmanlar, eksik}
    notlar: list = field(default_factory=list)


# ----------------------------------------------------------------------------
# Yöntemler
# ----------------------------------------------------------------------------

class Motor:
    """Ortak eksi stok mantığı: eksik kısım son maliyetle maliyetlenir; sonraki giriş önce eksiği kapatır."""

    def __init__(self):
        self.smm = SIFIR
        self.son_maliyet: Decimal | None = None
        self.eksik_miktar = SIFIR
        self.eksik_deger = SIFIR
        self.duzeltme = SIFIR
        self.son_cikis_birim: Decimal | None = None
        self.eksi_olay = []

    def giris(self, q: Decimal, p: Decimal) -> None:
        if self.eksik_miktar > 0:
            kapat = min(q, self.eksik_miktar)
            birim = self.eksik_deger / self.eksik_miktar
            fark = kapat * (p - birim)
            self.smm += fark
            self.duzeltme += fark
            self.eksik_deger -= birim * kapat
            self.eksik_miktar -= kapat
            q -= kapat
        if q > 0:
            self._ekle(q, p)
        self.son_maliyet = p

    def cikis(self, q: Decimal, t: date) -> Decimal:
        mevcut = self.miktar
        normal = min(q, max(mevcut, SIFIR))
        maliyet = self._cikar(normal) if normal > 0 else SIFIR
        if q > normal:
            eksik = q - normal
            birim = self.son_maliyet if self.son_maliyet is not None else SIFIR
            maliyet += eksik * birim
            self.eksik_miktar += eksik
            self.eksik_deger += eksik * birim
            self.eksi_olay.append((t, eksik, birim))
        self.smm += maliyet
        if q:
            self.son_cikis_birim = maliyet / q
        return maliyet

    def alis_iade(self, q: Decimal, p: Decimal) -> None:
        self._iade(q, p)

    def satis_iade(self, q: Decimal) -> Decimal:
        p = self.son_cikis_birim if self.son_cikis_birim is not None else (self.ortalama or self.son_maliyet or SIFIR)
        self._ekle(q, p)
        self.smm -= q * p
        return q * p


class Fifo(Motor):
    def __init__(self):
        super().__init__()
        self.katmanlar: list[list] = []      # [miktar, birim, tarih]
        self.t: date | None = None

    @property
    def miktar(self) -> Decimal:
        return sum((k[0] for k in self.katmanlar), SIFIR)

    @property
    def deger(self) -> Decimal:
        return sum((k[0] * k[1] for k in self.katmanlar), SIFIR)

    @property
    def ortalama(self) -> Decimal | None:
        return self.deger / self.miktar if self.miktar else None

    def _ekle(self, q, p):
        self.katmanlar.append([q, p, self.t])

    def _cikar(self, q):
        maliyet = SIFIR
        while q > 0 and self.katmanlar:
            k = self.katmanlar[0]
            al = min(q, k[0])
            maliyet += al * k[1]
            k[0] -= al
            q -= al
            if k[0] == 0:
                self.katmanlar.pop(0)
        return maliyet

    def _iade(self, q, p):
        for k in reversed(self.katmanlar):          # önce aynı fiyatlı en yeni katman
            if k[1] == p and q > 0:
                al = min(q, k[0])
                k[0] -= al
                q -= al
        while q > 0 and self.katmanlar:              # kalan: en yeni katmanlardan, katman maliyetiyle
            k = self.katmanlar[-1]
            al = min(q, k[0])
            k[0] -= al
            q -= al
            if k[0] == 0:
                self.katmanlar.pop()
        self.katmanlar = [k for k in self.katmanlar if k[0] > 0]


class Hareketli(Motor):
    def __init__(self):
        super().__init__()
        self._miktar = SIFIR
        self._deger = SIFIR
        self.t = None

    @property
    def miktar(self):
        return self._miktar

    @property
    def deger(self):
        return self._deger

    @property
    def ortalama(self):
        return self._deger / self._miktar if self._miktar else None

    def _ekle(self, q, p):
        self._miktar += q
        self._deger += q * p

    def _cikar(self, q):
        m = yuv(q * self.ortalama) if q < self._miktar else self._deger
        self._miktar -= q
        self._deger -= m
        return m

    def _iade(self, q, p):
        self._miktar -= q
        self._deger -= q * p


def degerle(s: Stok, tarih_: date | None) -> list[dict]:
    uy = []
    hs = sorted([h for h in s.hareketler if not tarih_ or h.tarih <= tarih_], key=lambda h: (h.tarih, h.tur not in ("devir", "giris"), h.sira))
    for ad, sinif in (("fifo", Fifo), ("hareketli", Hareketli)):
        m = sinif()
        detay = []
        for h in hs:
            m.t = h.tarih
            maliyet = None
            if h.tur in ("devir", "giris"):
                m.giris(h.miktar, h.fiyat)
                maliyet = h.miktar * h.fiyat
            elif h.tur == "cikis":
                maliyet = m.cikis(h.miktar, h.tarih)
            elif h.tur == "alis_iade":
                m.alis_iade(h.miktar, h.fiyat)
                maliyet = h.miktar * h.fiyat
            elif h.tur == "satis_iade":
                maliyet = m.satis_iade(h.miktar)
            detay.append((h, maliyet, m.miktar - m.eksik_miktar, m.deger - m.eksik_deger, m.ortalama))
        s.sonuc[ad] = {"miktar": m.miktar - m.eksik_miktar, "deger": m.deger - m.eksik_deger, "smm": m.smm, "detay": detay,
                       "katmanlar": [list(k) for k in getattr(m, "katmanlar", [])], "duzeltme": m.duzeltme, "eksi": m.eksi_olay}
        if ad == "fifo":
            for t, q, p in m.eksi_olay:
                uy.append({"onem": "Yüksek", "tur": "Eksi stok", "stok": s.kod, "aciklama": f"{t:%d.%m.%Y}: çıkış mevcut stoktan {q:g} {s.birim} fazla; "
                           f"eksik kısım son maliyetle ({tl(p)}) maliyetlendi. Giriş kaydı gecikmiş olabilir"})
            if m.eksik_miktar > 0:
                uy.append({"onem": "Yüksek", "tur": "Dönem sonu eksi stok", "stok": s.kod, "aciklama": f"Dönem sonunda {m.eksik_miktar:g} {s.birim} eksi stok"})
    # Dönem sonu ağırlıklı ortalama
    gir_q = sum((h.miktar for h in hs if h.tur in ("devir", "giris")), SIFIR) - sum((h.miktar for h in hs if h.tur == "alis_iade"), SIFIR)
    gir_v = sum((h.miktar * h.fiyat for h in hs if h.tur in ("devir", "giris")), SIFIR) - sum((h.miktar * h.fiyat for h in hs if h.tur == "alis_iade"), SIFIR)
    cik_q = sum((h.miktar for h in hs if h.tur == "cikis"), SIFIR) - sum((h.miktar for h in hs if h.tur == "satis_iade"), SIFIR)
    ort = gir_v / gir_q if gir_q else SIFIR
    kalan = gir_q - cik_q
    s.sonuc["donemsel"] = {"miktar": kalan, "deger": yuv(kalan * ort) if kalan != 0 else SIFIR, "smm": gir_v - (yuv(kalan * ort) if kalan != 0 else SIFIR),
                           "ortalama": ort, "detay": [], "katmanlar": [], "duzeltme": SIFIR, "eksi": []}
    # Diğer kontroller
    alislar = [h for h in hs if h.tur == "giris"]
    if alislar and ort:
        son = alislar[-1]
        if abs(son.fiyat - ort) / ort > Decimal("0.25"):
            uy.append({"onem": "Orta", "tur": "Fiyat sapması", "stok": s.kod, "aciklama": f"Son alış fiyatı {tl(son.fiyat)}, dönem ortalaması {tl(ort)} "
                       f"(%{abs(son.fiyat - ort) / ort * 100:.0f} fark); birim veya fiyat hatası olabilir"})
    if hs and all(h.tur == "devir" for h in hs) and kalan:
        uy.append({"onem": "Orta", "tur": "Hareketsiz stok", "stok": s.kod, "aciklama": f"Dönemde hiç hareket yok ({kalan:g} {s.birim} devirden kalan); "
                   "yavaş dönen / atıl stok olabilir"})
    for h in hs:
        if h.tur == "satis_iade" and "iade" not in katla(h.belge) and h.fiyat is None:
            uy.append({"onem": "Bilgi", "tur": "Fiyatsız giriş", "stok": s.kod, "aciklama": f"{h.tarih:%d.%m.%Y} {h.belge}: fiyatı olmayan giriş satış iadesi "
                       "sayıldı (son çıkış maliyetiyle)"})
    return uy


def ngd_testi(stoklar: list[Stok], fiyatlar: dict, yontem: str) -> tuple[list[dict], list[dict]]:
    satirlar, uy = [], []
    for s in stoklar:
        f = fiyatlar.get(katla(s.kod))
        r = s.sonuc[yontem]
        if not f or r["miktar"] <= 0:
            continue
        birim = r["deger"] / r["miktar"]
        ngd = f[0] - f[1]
        dusuk = (birim - ngd) * r["miktar"] if ngd < birim else SIFIR
        satirlar.append({"s": s, "birim": birim, "fiyat": f[0], "gider": f[1], "ngd": ngd, "dusukluk": yuv(dusuk)})
        if dusuk > 0:
            uy.append({"onem": "Orta", "tur": "Değer düşüklüğü", "stok": s.kod, "aciklama": f"Net gerçekleşebilir değer {tl(ngd)} < birim maliyet {tl(birim)}; "
                       f"olası değer düşüklüğü {tl(dusuk)} TL"})
    return satirlar, uy


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
RENK = {"Yüksek": "FDE2E1", "Orta": "FFF4CE", "Bilgi": "E8F0FE"}
UST = Alignment(vertical="top", wrap_text=True)
PF = "#,##0.00"
PB = "#,##0.0000"
TUR_AD = {"devir": "Devir", "giris": "Giriş", "cikis": "Çıkış", "alis_iade": "Alış iadesi", "satis_iade": "Satış iadesi"}


def _baslik(ws, basliklar, genislik=()):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def rapor_yaz(cikti: Path, s: dict) -> None:
    yontem = s["yontem"]
    stoklar = s["stoklar"]
    wb = Workbook()
    oz = wb.active
    oz.title = "Değerleme Özeti"
    _baslik(oz, ["Stok Kodu", "Stok Adı", "Birim", "Dönem Sonu Miktar", "Birim Maliyet", f"Stok Değeri ({YONTEMLER[yontem]})", "SMM / Kullanılan",
                 "FIFO Değeri", "Hareketli Ort. Değeri", "Dönem Sonu Ort. Değeri", "Not"], (11, 26, 7, 12, 12, 16, 15, 14, 14, 14, 30))
    for st in stoklar:
        r = st.sonuc[yontem]
        oz.append([st.kod, st.ad, st.birim, float(r["miktar"]), float(r["deger"] / r["miktar"]) if r["miktar"] else None, float(r["deger"]), float(r["smm"]),
                   float(st.sonuc["fifo"]["deger"]), float(st.sonuc["hareketli"]["deger"]), float(st.sonuc["donemsel"]["deger"]),
                   "; ".join(sorted({u["tur"] for u in s["uyarilar"] if u["stok"] == st.kod}))])
        r_ = oz.max_row
        oz.cell(r_, 5).number_format = PB
        for j in (6, 7, 8, 9, 10):
            oz.cell(r_, j).number_format = PF
    oz.append(["Toplam", "", "", None, None] + [float(sum((st.sonuc[y][k] for st in stoklar), SIFIR)) for y, k in
                                                ((yontem, "deger"), (yontem, "smm"), ("fifo", "deger"), ("hareketli", "deger"), ("donemsel", "deger"))])
    for j in (6, 7, 8, 9, 10):
        oz.cell(oz.max_row, j).number_format = PF
        oz.cell(oz.max_row, j).font = Font(bold=True)

    yk = wb.create_sheet("Yöntem Karşılaştırması")
    _baslik(yk, ["Yöntem", "Dönem Sonu Stok Değeri", "SMM / Kullanılan", "Eksi Stok Düzeltmesi", "Fark (ana yönteme göre, stok)"], (32, 18, 18, 16, 18))
    ana = sum((st.sonuc[yontem]["deger"] for st in stoklar), SIFIR)
    for y, ad in YONTEMLER.items():
        d = sum((st.sonuc[y]["deger"] for st in stoklar), SIFIR)
        yk.append([ad + (" (ana)" if y == yontem else ""), float(d), float(sum((st.sonuc[y]["smm"] for st in stoklar), SIFIR)),
                   float(sum((st.sonuc[y]["duzeltme"] for st in stoklar), SIFIR)), float(d - ana)])
        for j in (2, 3, 4, 5):
            yk.cell(yk.max_row, j).number_format = PF
    yk.append([])
    yk.append(["Fiyatlar yükselirken FIFO stoku daha yüksek, SMM'yi daha düşük gösterir. Yöntem seçildikten sonra tutarlı uygulanmalıdır."])

    if yontem != "donemsel":
        hd = wb.create_sheet("Hareket Detayı")
        _baslik(hd, ["Stok Kodu", "Tarih", "Belge", "Hareket", "Miktar", "Birim Fiyat", "Hareket Maliyeti", "Kalan Miktar", "Kalan Değer", "Ortalama"],
                (11, 11, 15, 12, 10, 12, 14, 11, 14, 12))
        for st in stoklar:
            for h, mal, km, kd, ort in st.sonuc[yontem]["detay"]:
                hd.append([st.kod, h.tarih, h.belge, TUR_AD[h.tur], float(h.miktar), None if h.fiyat is None else float(h.fiyat),
                           None if mal is None else float(yuv(mal)), float(km), float(yuv(kd)), None if ort is None else float(ort)])
                r_ = hd.max_row
                hd.cell(r_, 2).number_format = "DD.MM.YYYY"
                for j in (6, 7, 9):
                    hd.cell(r_, j).number_format = PF
                hd.cell(r_, 10).number_format = PB
        hd.auto_filter.ref = f"A1:J{hd.max_row}"

    fk = wb.create_sheet("FIFO Katmanları")
    _baslik(fk, ["Stok Kodu", "Stok Adı", "Giriş Tarihi", "Kalan Miktar", "Birim Maliyet", "Değer"], (11, 26, 12, 12, 12, 14))
    for st in stoklar:
        for q, p, t in st.sonuc["fifo"]["katmanlar"]:
            fk.append([st.kod, st.ad, t, float(q), float(p), float(yuv(q * p))])
            fk.cell(fk.max_row, 3).number_format = "DD.MM.YYYY"
            fk.cell(fk.max_row, 5).number_format = PB
            fk.cell(fk.max_row, 6).number_format = PF

    if s["ngd"]:
        nd = wb.create_sheet("Değer Düşüklüğü")
        _baslik(nd, ["Stok Kodu", "Stok Adı", "Miktar", "Birim Maliyet", "Tahmini Satış Fiyatı", "Satış Gideri", "Net Gerçekleşebilir Değer",
                     "Olası Değer Düşüklüğü", "Karar"], (11, 26, 10, 12, 13, 11, 14, 14, 20))
        for x in s["ngd"]:
            st = x["s"]
            nd.append([st.kod, st.ad, float(st.sonuc[yontem]["miktar"]), float(x["birim"]), float(x["fiyat"]), float(x["gider"]), float(x["ngd"]),
                       float(x["dusukluk"]), ""])
            for j in (4, 5, 6, 7, 8):
                nd.cell(nd.max_row, j).number_format = PF
            if x["dusukluk"]:
                nd.cell(nd.max_row, 8).fill = PatternFill("solid", fgColor="FDE2E1")
            nd.cell(nd.max_row, 9).fill = PatternFill("solid", fgColor="FFF4CE")

    uy = wb.create_sheet("Uyarılar")
    _baslik(uy, ["Önem", "Tür", "Stok", "Açıklama"], (9, 20, 11, 100))
    for u in s["uyarilar"]:
        uy.append([u["onem"], u["tur"], u["stok"], u["aciklama"]])
        uy.cell(uy.max_row, 1).fill = PatternFill("solid", fgColor=RENK[u["onem"]])
        uy.cell(uy.max_row, 4).alignment = UST
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(hareket_yolu: Path, cikti: Path, yontem: str = "fifo", tarih_: date | None = None, fiyat_yolu: Path | None = None) -> dict:
    stoklar: dict[str, Stok] = {}
    uyarilar = []
    for r in kayitlar(hareket_yolu, SUTUNLAR, ("kod", "hareket", "miktar")):
        kod = metin(r.get("kod"))
        if not kod:
            continue
        q, p, t = sayi(r.get("miktar")), sayi(r.get("fiyat")), tarih(r.get("tarih"))
        if p is None and sayi(r.get("tutar")) is not None and q:
            p = sayi(r.get("tutar")) / q
        tur = tur_bul(r.get("hareket"), p)
        if q is None or t is None or tur is None or (tur in ("devir", "giris", "alis_iade") and p is None):
            uyarilar.append({"onem": "Orta", "tur": "Okunamayan satır", "stok": kod, "aciklama": f"Satır {r['_satir']}: tarih, miktar, hareket türü veya fiyat okunamadı"})
            continue
        st = stoklar.setdefault(kod, Stok(kod, metin(r.get("ad")), metin(r.get("birim"))))
        st.hareketler.append(Hareket(t, tur, abs(q), p, metin(r.get("belge")), r["_satir"]))
    if not stoklar:
        raise ValueError(f"{hareket_yolu.name}: stok hareketi bulunamadı")
    liste = sorted(stoklar.values(), key=lambda s: s.kod)
    for st in liste:
        uyarilar += degerle(st, tarih_)
    fiyatlar = {}
    if fiyat_yolu:
        for r in kayitlar(fiyat_yolu, FIYAT_SUTUNLARI, ("kod", "fiyat")):
            if metin(r.get("kod")) and sayi(r.get("fiyat")) is not None:
                fiyatlar[katla(r["kod"])] = (sayi(r["fiyat"]), sayi(r.get("gider")) or SIFIR)
    ngd, u2 = ngd_testi(liste, fiyatlar, yontem)
    uyarilar += u2
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    uyarilar.sort(key=lambda u: (sira[u["onem"]], u["tur"], u["stok"]))
    s = {"stoklar": liste, "uyarilar": uyarilar, "yontem": yontem, "ngd": ngd}
    rapor_yaz(cikti, s)
    return s


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Stok hareketlerinden FIFO / ağırlıklı ortalama ile dönem sonu stok değerini ve SMM'yi hesaplar.")
    p.add_argument("--hareketler", type=Path, default=ORNEK / "hareketler.csv",
                   help="Tarih, Stok Kodu, Stok Adı, Birim, Hareket (Devir/Giriş/Çıkış/Alış İadesi/Satış İadesi), Miktar, Birim Fiyat, Belge No")
    p.add_argument("--yontem", choices=list(YONTEMLER), default="fifo", help="Ana yöntem (varsayılan fifo); üçü de karşılaştırmada hesaplanır")
    p.add_argument("--tarih", help="Değerleme tarihi GG.AA.YYYY; sonraki hareketler dikkate alınmaz")
    p.add_argument("--satis-fiyatlari", type=Path, help="İsteğe bağlı: Stok Kodu, Tahmini Satış Fiyatı, Satış Gideri (net gerçekleşebilir değer testi)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "stok_degerleme.xlsx")
    a = p.parse_args(argv)
    fiyat = a.satis_fiyatlari or (ORNEK / "satis_fiyatlari.csv" if a.hareketler == ORNEK / "hareketler.csv" else None)
    for y in (a.hareketler, fiyat):
        if y and not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    t = tarih(a.tarih) if a.tarih else None
    if a.tarih and not t:
        print("[X] --tarih GG.AA.YYYY biçiminde olmalı")
        return 2
    try:
        s = calistir(a.hareketler, a.cikti, a.yontem, t, fiyat)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    for y, ad in YONTEMLER.items():
        print(f"[OK] {ad}: stok {tl(sum((st.sonuc[y]['deger'] for st in s['stoklar']), SIFIR))} TL · SMM "
              f"{tl(sum((st.sonuc[y]['smm'] for st in s['stoklar']), SIFIR))} TL" + ("  ← ana yöntem" if y == a.yontem else ""))
    for u in s["uyarilar"]:
        if u["onem"] != "Bilgi":
            print(f"[!] {u['stok']} {u['tur']}: {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
