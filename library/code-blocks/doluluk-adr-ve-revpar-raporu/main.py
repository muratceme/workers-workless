"""
Doluluk, ADR ve RevPAR Raporu — Workers / Workless kod bloğu
Otel › Ön Büro / Gelir Yönetimi

Oda satışlarından günlük ve aylık otel performans göstergelerini hesaplar ve geçen yılla karşılaştırır:

    Kullanılabilir oda = oda sayısı × gün (− kullanım dışı odalar, --ooo-dus verilirse)
    Doluluk  = ücretli satılan oda / kullanılabilir oda
    ADR      = oda geliri / ücretli satılan oda          (ortalama oda fiyatı)
    RevPAR   = oda geliri / kullanılabilir oda = Doluluk × ADR

Ücretsiz (complimentary) ve kendi kullanım (house use) odaları ADR'ye girmez; bunlarla birlikte "toplam doluluk"
ayrıca gösterilir. Geçen yıl karşılaştırması haftanın aynı gününe hizalanır (364 gün önce). Girdi, PMS'in günlük
istatistik raporu veya rezervasyon listesi (giriş/çıkış tarihi) olabilir; rezervasyonlar gecelere açılır ve
kanal/segment kırılımı çıkarılır. Oda geliri KDV ve konaklama vergisi hariç olmalıdır. İnternete bağlanmaz.

Kullanım:
    python main.py                                            # örnek rezervasyonlarla dener (120 oda)
    python main.py --girdi gunluk_istatistik.xlsx --oda 120
    python main.py --girdi rezervasyonlar.xlsx --oda 120 --baslangic 01.09.2026 --bitis 30.09.2026
"""
from __future__ import annotations

import argparse
import csv
import sys
from collections import OrderedDict, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]


def kucuk(s) -> str:
    return str(s or "").replace("İ", "i").replace("I", "ı").lower().strip()


def sayi(x) -> Decimal:
    if x in (None, ""):
        return Decimal(0)
    if isinstance(x, (int, float)):
        return Decimal(str(x))
    s = str(x).strip().replace("TL", "").replace("₺", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return Decimal(s)
    except InvalidOperation:
        return Decimal(0)


def tarih(x) -> date | None:
    if x in (None, ""):
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    for f in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y %H:%M"):
        try:
            return datetime.strptime(str(x).strip(), f).date()
        except ValueError:
            pass
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


ALANLAR = {
    "tarih": ("tarih", "gün", "konaklama tarihi"),
    "satilan": ("satılan oda", "dolu oda", "oda satışı", "ücretli oda"),
    "comp": ("ücretsiz oda", "comp", "complimentary", "ücretsiz"),
    "house": ("kendi kullanım", "house use", "personel kullanımı"),
    "ooo": ("kullanım dışı", "arızalı oda", "ooo", "out of order"),
    "toplam_oda": ("toplam oda", "oda sayısı", "kapasite"),
    "gelir": ("oda geliri", "net oda geliri", "konaklama geliri", "gelir"),
    "giris": ("giriş tarihi", "giriş", "check-in", "geliş tarihi"),
    "cikis": ("çıkış tarihi", "çıkış", "check-out", "ayrılış tarihi"),
    "oda_adedi": ("oda adedi", "oda sayısı (rez.)", "oda"),
    "gecelik": ("gecelik fiyat", "gecelik oda fiyatı", "oda fiyatı"),
    "toplam_tutar": ("toplam tutar", "konaklama tutarı", "rezervasyon tutarı"),
    "kanal": ("kanal", "segment", "pazar", "acente"),
    "durum": ("durum", "statü", "rezervasyon durumu"),
    "ucret_tipi": ("ücret tipi", "fiyat tipi", "rate code"),
}


def sutunlar(baslik):
    b = [kucuk(x) for x in baslik]
    return {alan: next((i for i, x in enumerate(b) if x in adlar), None) for alan, adlar in ALANLAR.items()}


def gunluk_veri(yol: Path, oda: int | None) -> tuple[dict, dict, list[str]]:
    """Dönüş: (tarih → {satilan, comp, house, ooo, oda, gelir}, (tarih, kanal) → {satilan, gelir}, uyarılar)."""
    s = tablo_oku(yol)
    k = sutunlar(s[0])
    al = lambda r, a: r[k[a]] if k[a] is not None and k[a] < len(r) else None  # noqa: E731
    gun: dict[date, dict] = defaultdict(lambda: defaultdict(Decimal))
    kanal: dict[tuple, dict] = defaultdict(lambda: defaultdict(Decimal))
    uyarilar = []
    if k["giris"] is not None and k["cikis"] is not None:            # rezervasyon listesi
        iptal = 0
        for r in s[1:]:
            g, c = tarih(al(r, "giris")), tarih(al(r, "cikis"))
            if not g or not c or c <= g:
                continue
            d = kucuk(al(r, "durum"))
            if "iptal" in d or "no show" in d or "no-show" in d or "gelmedi" in d:
                iptal += 1
                continue
            adet = sayi(al(r, "oda_adedi")) or Decimal(1)
            gece = (c - g).days
            if k["gecelik"] is not None and al(r, "gecelik") not in (None, ""):
                gecelik = sayi(al(r, "gecelik")) * adet
            else:
                gecelik = sayi(al(r, "toplam_tutar")) / gece
            tip = kucuk(al(r, "ucret_tipi"))
            tur = "comp" if ("ücretsiz" in tip or "comp" in tip) else "house" if ("house" in tip or "kendi" in tip) else "satilan"
            kn = str(al(r, "kanal") or "Belirtilmemiş").strip()
            for i in range(gece):
                t = g + timedelta(days=i)
                gun[t][tur] += adet
                if tur == "satilan":
                    gun[t]["gelir"] += gecelik
                    kanal[(t, kn)]["satilan"] += adet
                    kanal[(t, kn)]["gelir"] += gecelik
        if iptal:
            uyarilar.append(f"{iptal} iptal/no-show rezervasyon hesaba katılmadı (no-show geliri varsa günlük rapora ekleyin)")
    else:
        if k["tarih"] is None or k["satilan"] is None or k["gelir"] is None:
            raise SystemExit("Günlük raporda Tarih, Satılan Oda ve Oda Geliri; rezervasyon listesinde Giriş ve Çıkış Tarihi gerekli. "
                             f"Başlıklar: {s[0]}")
        for r in s[1:]:
            t = tarih(al(r, "tarih"))
            if not t:
                continue
            for alan in ("satilan", "comp", "house", "ooo", "gelir", "toplam_oda"):
                if k[alan] is not None:
                    gun[t][alan] += sayi(al(r, alan))
    for t, v in gun.items():
        if not v.get("toplam_oda"):
            if not oda:
                raise SystemExit("Toplam oda sayısını --oda ile verin (veya günlük raporda 'Toplam Oda' sütunu olsun).")
            v["toplam_oda"] = Decimal(oda)
        dolu = v["satilan"] + v["comp"] + v["house"]
        if dolu > v["toplam_oda"]:
            uyarilar.append(f"{t:%d.%m.%Y}: dolu oda ({dolu}) oda sayısını ({v['toplam_oda']}) aşıyor — overbooking veya veri hatası")
    return gun, kanal, uyarilar


# ----------------------------------------------------------------------------
# Göstergeler
# ----------------------------------------------------------------------------

def gostergeler(v: dict, ooo_dus: bool) -> dict:
    kullanilabilir = v["toplam_oda"] - (v["ooo"] if ooo_dus else 0)
    satilan, gelir = v["satilan"], v["gelir"]
    return {
        "kullanilabilir": kullanilabilir, "satilan": satilan, "comp": v["comp"], "house": v["house"], "ooo": v["ooo"], "gelir": gelir,
        "doluluk": float(satilan / kullanilabilir) if kullanilabilir else None,
        "toplam_doluluk": float((satilan + v["comp"] + v["house"]) / kullanilabilir) if kullanilabilir else None,
        "adr": float(gelir / satilan) if satilan else None,
        "revpar": float(gelir / kullanilabilir) if kullanilabilir else None,
    }


def topla(gunler: list[dict]) -> dict:
    t = defaultdict(Decimal)
    for v in gunler:
        for k in ("toplam_oda", "satilan", "comp", "house", "ooo", "gelir"):
            t[k] += v.get(k, Decimal(0))
    return t


def degisim(a, b):
    return None if a is None or b in (None, 0) else a / b - 1


def calistir(girdi: Path, cikti: Path, oda: int | None = None, baslangic: date | None = None, bitis: date | None = None,
             ooo_dus: bool = False) -> dict:
    gun, kanal, uyarilar = gunluk_veri(girdi, oda)
    if not gun:
        raise SystemExit("Hesaplanacak gün yok.")
    bas = baslangic or min(gun)
    bit = bitis or max(gun)
    donem = [t for t in sorted(gun) if bas <= t <= bit]
    gunluk = OrderedDict()
    for t in donem:
        g = gostergeler(gun[t], ooo_dus)
        ly_t = t - timedelta(days=364)                       # haftanın aynı günü
        g["ly"] = gostergeler(gun[ly_t], ooo_dus) if ly_t in gun else None
        gunluk[t] = g
    aylik = OrderedDict()
    for t in donem:
        aylik.setdefault((t.year, t.month), []).append(t)
    aylar = OrderedDict()
    for (y, m), ts in aylik.items():
        a = gostergeler(topla([gun[t] for t in ts]), ooo_dus)
        ly = [t - timedelta(days=364) for t in ts]
        a["ly"] = gostergeler(topla([gun[t] for t in ly]), ooo_dus) if all(t in gun for t in ly) else None
        a["gun_sayisi"] = len(ts)
        aylar[(y, m)] = a
    hafta = OrderedDict((i, gostergeler(topla([gun[t] for t in donem if t.weekday() == i]), ooo_dus)) for i in range(7))
    kanallar = defaultdict(lambda: defaultdict(Decimal))
    for (t, kn), v in kanal.items():
        if bas <= t <= bit:
            kanallar[kn]["satilan"] += v["satilan"]
            kanallar[kn]["gelir"] += v["gelir"]
    toplam = gostergeler(topla([gun[t] for t in donem]), ooo_dus)
    _rapor(gunluk, aylar, hafta, kanallar, toplam, cikti, uyarilar, ooo_dus)
    return {"gunluk": gunluk, "aylar": aylar, "hafta": hafta, "kanallar": kanallar, "toplam": toplam, "uyarilar": uyarilar}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
YESIL = PatternFill("solid", fgColor="E3F5E1")
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
PARA = "#,##0.00"


def f(x):
    return None if x is None else float(x)


def _baslik(ws, basliklar, genislik):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w


def _bicim(ws, para_s=(), yuzde_s=(), degisim_s=()):
    r = ws.max_row
    for c in para_s:
        ws.cell(r, c).number_format = PARA
    for c in yuzde_s:
        ws.cell(r, c).number_format = "0.0%"
    for c in degisim_s:
        ws.cell(r, c).number_format = "+0.0%;-0.0%;0.0%"
        v = ws.cell(r, c).value
        if v is not None:
            ws.cell(r, c).fill = YESIL if v > 0 else KIRMIZI if v < 0 else PatternFill()


def _rapor(gunluk, aylar, hafta, kanallar, toplam, cikti, uyarilar, ooo_dus):
    wb = Workbook()
    a = wb.active
    a.title = "Aylık"
    _baslik(a, ["Ay", "Gün", "Kullanılabilir Oda", "Satılan (ücretli)", "Ücretsiz + Kendi K.", "Oda Geliri", "Doluluk", "Toplam Doluluk",
                "ADR", "RevPAR", "GY Doluluk", "GY ADR", "GY RevPAR", "Doluluk Δ (puan)", "ADR Δ", "RevPAR Δ"],
            (14, 6, 14, 13, 14, 16, 10, 12, 12, 12, 11, 12, 12, 13, 9, 10))
    for (y, m), x in aylar.items():
        ly = x["ly"]
        a.append([f"{AYLAR[m - 1]} {y}", x["gun_sayisi"], f(x["kullanilabilir"]), f(x["satilan"]), f(x["comp"] + x["house"]), f(x["gelir"]),
                  x["doluluk"], x["toplam_doluluk"], x["adr"], x["revpar"],
                  ly and ly["doluluk"], ly and ly["adr"], ly and ly["revpar"],
                  None if not ly or ly["doluluk"] is None else (x["doluluk"] - ly["doluluk"]),
                  degisim(x["adr"], ly and ly["adr"]), degisim(x["revpar"], ly and ly["revpar"])])
        _bicim(a, (6, 9, 10, 12, 13), (7, 8, 11), (15, 16))
        a.cell(a.max_row, 14).number_format = '+0.0%;-0.0%;0.0%'
    a.append(["DÖNEM TOPLAMI", None, f(toplam["kullanilabilir"]), f(toplam["satilan"]), f(toplam["comp"] + toplam["house"]),
              f(toplam["gelir"]), toplam["doluluk"], toplam["toplam_doluluk"], toplam["adr"], toplam["revpar"]])
    _bicim(a, (6, 9, 10), (7, 8))
    for h in a[a.max_row]:
        h.font = Font(bold=True)
    a.freeze_panes = "B2"
    a.append([])
    a.append(["GY: geçen yıl, haftanın aynı gününe hizalı (364 gün önce). Doluluk farkı yüzde puan olarak gösterilir."])
    for u in uyarilar[:50]:
        a.append(["Uyarı", u])

    g = wb.create_sheet("Günlük")
    _baslik(g, ["Tarih", "Gün", "Kullanılabilir", "Satılan", "Ücretsiz", "Kendi K.", "Kullanım Dışı", "Oda Geliri", "Doluluk", "ADR",
                "RevPAR", "GY Tarih", "GY Doluluk", "GY ADR", "GY RevPAR", "RevPAR Δ"],
            (12, 11, 12, 9, 9, 9, 11, 14, 9, 11, 11, 12, 11, 11, 11, 10))
    for t, x in gunluk.items():
        ly = x["ly"]
        g.append([t, GUNLER[t.weekday()], f(x["kullanilabilir"]), f(x["satilan"]), f(x["comp"]), f(x["house"]), f(x["ooo"]), f(x["gelir"]),
                  x["doluluk"], x["adr"], x["revpar"], t - timedelta(days=364) if ly else None, ly and ly["doluluk"], ly and ly["adr"],
                  ly and ly["revpar"], degisim(x["revpar"], ly and ly["revpar"])])
        g.cell(g.max_row, 1).number_format = g.cell(g.max_row, 12).number_format = "DD.MM.YYYY"
        _bicim(g, (8, 10, 11, 14, 15), (9, 13), (16,))
    g.freeze_panes = "B2"
    n = len(gunluk)
    if n:
        lc = LineChart()
        lc.title, lc.height, lc.width = "Günlük doluluk", 8, 24
        lc.add_data(Reference(g, min_col=9, min_row=1, max_row=1 + n), titles_from_data=True)
        lc.add_data(Reference(g, min_col=13, min_row=1, max_row=1 + n), titles_from_data=True)
        lc.set_categories(Reference(g, min_col=1, min_row=2, max_row=1 + n))
        lc.y_axis.number_format = "0%"
        g.add_chart(lc, "R2")
        bc = BarChart()
        bc.title, bc.height, bc.width = "Günlük ADR ve RevPAR", 8, 24
        bc.add_data(Reference(g, min_col=10, max_col=11, min_row=1, max_row=1 + n), titles_from_data=True)
        bc.set_categories(Reference(g, min_col=1, min_row=2, max_row=1 + n))
        g.add_chart(bc, "R20")

    h = wb.create_sheet("Haftanın Günü")
    _baslik(h, ["Gün", "Kullanılabilir", "Satılan", "Oda Geliri", "Doluluk", "ADR", "RevPAR"], (12, 13, 10, 15, 10, 11, 11))
    for i, x in hafta.items():
        h.append([GUNLER[i], f(x["kullanilabilir"]), f(x["satilan"]), f(x["gelir"]), x["doluluk"], x["adr"], x["revpar"]])
        _bicim(h, (4, 6, 7), (5,))

    if kanallar:
        k = wb.create_sheet("Kanal")
        _baslik(k, ["Kanal / Segment", "Oda Gecesi", "Pay", "Oda Geliri", "Gelir Payı", "ADR"], (24, 12, 9, 16, 10, 12))
        ts = sum((v["satilan"] for v in kanallar.values()), Decimal(0))
        tg = sum((v["gelir"] for v in kanallar.values()), Decimal(0))
        for ad, v in sorted(kanallar.items(), key=lambda i: -i[1]["gelir"]):
            k.append([ad, f(v["satilan"]), f(v["satilan"] / ts) if ts else None, f(v["gelir"]), f(v["gelir"] / tg) if tg else None,
                      f(v["gelir"] / v["satilan"]) if v["satilan"] else None])
            _bicim(k, (4, 6), (3, 5))

    b = wb.create_sheet("Bilgi")
    for s in [["Kullanılabilir oda", "Oda sayısı × gün" + (" − kullanım dışı (OOO) odalar" if ooo_dus else
                                                           "; kullanım dışı odalar düşülmez (düşmek için --ooo-dus)")],
              ["Doluluk", "Ücretli satılan oda / kullanılabilir oda"],
              ["Toplam doluluk", "(Ücretli + ücretsiz + kendi kullanım) / kullanılabilir oda"],
              ["ADR", "Oda geliri / ücretli satılan oda (ücretsiz ve kendi kullanım hariç)"],
              ["RevPAR", "Oda geliri / kullanılabilir oda = doluluk × ADR"],
              ["Oda geliri", "KDV ve konaklama vergisi hariç; paket fiyatlarda (kahvaltı, yarım pansiyon vb.) yalnız oda payı"],
              ["Geçen yıl", "364 gün önceki gün (haftanın aynı günü); geçen yıl verisi dosyada yoksa boş kalır"]]:
        b.append(s)
    b.column_dimensions["A"].width = 18
    b.column_dimensions["B"].width = 120
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="Doluluk, ADR ve RevPAR'ı günlük/aylık ve geçen yıl karşılaştırmalı hesaplar.")
    ap.add_argument("--girdi", type=Path, default=BURASI / "ornek_veri" / "rezervasyonlar.csv",
                    help="Günlük istatistik (Tarih, Satılan Oda, Oda Geliri, ...) veya rezervasyon listesi (Giriş, Çıkış, Gecelik Fiyat, ...)")
    ap.add_argument("--oda", type=int, help="Toplam oda sayısı (günlük raporda 'Toplam Oda' yoksa)")
    ap.add_argument("--baslangic", help="Rapor başlangıcı (GG.AA.YYYY)")
    ap.add_argument("--bitis", help="Rapor bitişi (GG.AA.YYYY)")
    ap.add_argument("--ooo-dus", action="store_true", help="Kullanım dışı (arızalı) odaları kullanılabilir odadan düş")
    ap.add_argument("--cikti", type=Path, default=Path("cikti") / "doluluk_adr_revpar.xlsx")
    a = ap.parse_args(argv)
    if a.girdi == BURASI / "ornek_veri" / "rezervasyonlar.csv":
        a.oda = a.oda or 40
        a.baslangic = a.baslangic or "01.09.2026"
        a.bitis = a.bitis or "30.09.2026"
    s = calistir(a.girdi, a.cikti, a.oda, tarih(a.baslangic), tarih(a.bitis), a.ooo_dus)
    t = s["toplam"]
    yz = lambda x: "—" if x is None else f"%{x * 100:.1f}".replace(".", ",")  # noqa: E731
    tl = lambda x: "—" if x is None else f"{x:,.2f} TL".replace(",", "X").replace(".", ",").replace("X", ".")  # noqa: E731
    print(f"[OK] {len(s['gunluk'])} gün · doluluk {yz(t['doluluk'])} · ADR {tl(t['adr'])} · RevPAR {tl(t['revpar'])}")
    for (y, m), x in s["aylar"].items():
        ly = x["ly"]
        print(f"     {AYLAR[m - 1]} {y}: doluluk {yz(x['doluluk'])}, ADR {tl(x['adr'])}, RevPAR {tl(x['revpar'])}"
              + (f" · GY RevPAR {tl(ly['revpar'])}" if ly else ""))
    for u in s["uyarilar"][:10]:
        print(f"[!] {u}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")


if __name__ == "__main__":
    sys.exit(main())
