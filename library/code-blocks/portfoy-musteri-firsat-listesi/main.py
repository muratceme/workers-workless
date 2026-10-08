"""
Portföy Müşteri Fırsat Listesi — Workers / Workless kod bloğu
Bankacılık › Şube Bankacılığı › Bireysel Portföy Yöneticisi

Portföydeki bireysel müşterilerin ürün sahipliği ve hareketlerinden çapraz satış ve elde tutma fırsatlarını
kural tabanlı olarak puanlar ve önceliklendirir:
  - Fırsatlar: kredi kartı, BES, vadeli mevduat / yatırım ürünü (yüksek vadesiz bakiye), vade yenileme (vadesi
    yaklaşan mevduat), kasko (taşıt kredisi var), konut sigortası (konut kredisi var), hayat sigortası, mobil
    bankacılık, otomatik ödeme talimatı, pasif müşteri reaktivasyonu, kredisi kapanan müşteriyle ihtiyaç görüşmesi.
  - Sorumlu satış: gecikmesi olan veya kart limit kullanımı yüksek (--limit-esik) müşteriye kart/kredi önerilmez;
    sigortalar krediye bağlı zorunlu ürün olarak sunulmaz (not düşülür).
  - İletişim izni: pazarlama (İYS) izni olmayan müşteriler için uzaktan ticari ileti kanalı önerilmez, yalnız
    şube görüşmesi notu düşülür; --izinsiz-haric ile listeden çıkarılır.
  - Eşikler ve puanlar ayar dosyasından (JSON) değiştirilebilir.
İnternete bağlanmaz.

Kullanım:
    python main.py                                                 # örnek portföy (rapor tarihi 08.10.2026)
    python main.py --girdi portfoy.xlsx --ayarlar kurallar.json --tarih 08.10.2026
    python main.py --girdi portfoy.xlsx --izinsiz-haric --ilk 3
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri" / "portfoy.csv"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")

VARSAYILAN = {
    "min_gelir_kart": 25000, "bes_yas": [18, 56], "min_gelir_bes": 20000, "vadesiz_esik": 150000, "vadesiz_yuksek": 500000,
    "vade_yenileme_gun": 15, "pasif_gun": 90, "kredi_bitis_gun": 60, "limit_esik": 90, "hayat_yas": [25, 60],
    "puan": {"kart": 30, "bes": 25, "mevduat": 30, "vade": 40, "kasko": 25, "konut_sig": 20, "hayat": 15, "mobil": 15, "talimat": 15,
             "reaktivasyon": 20, "kredi_bitis": 10, "maas_bonus": 10, "yuksek_bakiye_bonus": 10},
}
URUNLER = {"kart": "Kredi kartı", "bes": "Bireysel emeklilik (BES)", "mevduat": "Vadeli mevduat / yatırım fonu", "vade": "Vade yenileme",
           "kasko": "Kasko", "konut_sig": "Konut sigortası", "hayat": "Hayat sigortası", "mobil": "Mobil bankacılık", "talimat": "Otomatik ödeme talimatı",
           "reaktivasyon": "Reaktivasyon görüşmesi", "kredi_bitis": "İhtiyaç görüşmesi (kredi kapanıyor)"}
SUTUNLAR = {
    "no": ("musteri no", "musteri"), "ad": ("ad soyad", "musteri adi", "ad"), "yas": ("yas",), "maas": ("maas musterisi", "maas"),
    "gelir": ("aylik gelir", "gelir", "maas tutari"), "vadesiz": ("vadesiz ort. bakiye (3 ay)", "vadesiz ortalama bakiye", "vadesiz bakiye", "vadesiz"),
    "vadeli": ("vadeli mevduat", "vadeli bakiye", "vadeli"), "vade_tarihi": ("vadeli vade tarihi", "vade tarihi", "mevduat vade tarihi"),
    "fon": ("yatirim fonu", "fon bakiyesi", "yatirim"), "kart": ("kredi karti", "kart"), "limit": ("kart limit kullanimi (%)", "limit kullanimi", "kart limit kullanimi"),
    "bes": ("bes",), "hayat": ("hayat sigortasi",), "konut": ("konut kredisi",), "konut_sig": ("konut sigortasi",), "tasit": ("tasit kredisi",),
    "kasko": ("kasko",), "ihtiyac": ("ihtiyac kredisi",), "kredi_bitis": ("kredi bitis tarihi", "kredi vade sonu"),
    "mobil": ("mobil bankacilik", "dijital", "mobil"), "talimat": ("otomatik odeme talimati", "fatura talimati", "talimat sayisi"),
    "son_islem": ("son islem tarihi", "son islem"), "gecikme": ("gecikme", "gecikmesi var", "takip"), "izin": ("pazarlama izni", "iys izni", "ticari ileti izni"),
}


def katla(s) -> str:
    return " ".join(str(s or "").translate(_TR).lower().split())


def sayi(x) -> float:
    if x in (None, ""):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = re.sub(r"[^\d,.\-]", "", str(x))
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[1]) == 3):
        s = s.replace(".", "")
    try:
        return float(Decimal(s))
    except InvalidOperation:
        return 0.0


def evet(x) -> bool:
    k = katla(x)
    if k in ("e", "evet", "var", "x", "1", "true", "y", "yes"):
        return True
    return sayi(x) > 0 if re.search(r"\d", k) else False


def tarih(x) -> date | None:
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    for f in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(x or "").strip(), f).date()
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


@dataclass
class Musteri:
    ham: dict
    firsatlar: list[tuple[str, int, str]] = field(default_factory=list)     # (ürün kodu, puan, gerekçe)
    notlar: list[str] = field(default_factory=list)
    kanal: str = ""

    def __getattr__(self, a):
        if a in SUTUNLAR:
            return self.ham.get(a)
        raise AttributeError(a)

    @property
    def toplam(self) -> int:
        return sum(p for _, p, _ in sorted(self.firsatlar, key=lambda f: -f[1])[:3])


def portfoy_oku(yol: Path) -> tuple[list[Musteri], list[str]]:
    s = tablo_oku(yol)
    for bi, r in enumerate(s[:10]):
        b = [katla(x) for x in r]
        k = {a: next((i for i, x in enumerate(b) if x in es), None) for a, es in SUTUNLAR.items()}
        if k["no"] is not None:
            break
    else:
        raise SystemExit(f"{yol.name}: 'Müşteri No' sütunu bulunamadı.")
    eksik = [a for a, i in k.items() if i is None]
    return [Musteri({a: (r[i] if i is not None and i < len(r) else None) for a, i in k.items()}) for r in s[bi + 1:] if r[k["no"]]], eksik


URUN_SAHIPLIGI = [("Kredi kartı", "kart"), ("BES", "bes"), ("Vadeli mevduat", "vadeli"), ("Yatırım fonu", "fon"), ("Hayat sigortası", "hayat"),
                  ("Konut kredisi", "konut"), ("Konut sigortası", "konut_sig"), ("Taşıt kredisi", "tasit"), ("Kasko", "kasko"),
                  ("İhtiyaç kredisi", "ihtiyac"), ("Mobil bankacılık", "mobil"), ("Otomatik ödeme talimatı", "talimat"), ("Maaş müşterisi", "maas")]


def urun_sayisi(m: Musteri) -> int:
    return sum(1 for _, a in URUN_SAHIPLIGI if evet(m.ham.get(a)))


def degerlendir(m: Musteri, a: dict, bugun: date) -> None:
    p = a["puan"]
    yas, gelir, maas = int(sayi(m.yas)), sayi(m.gelir), evet(m.maas)
    gecikme, limit = evet(m.gecikme), sayi(m.limit)
    borclandirma_yok = gecikme or (evet(m.kart) and limit >= a["limit_esik"])
    if gecikme:
        m.notlar.append("Gecikmesi var: kart/kredi ürünü önerilmez")
    if evet(m.kart) and limit >= a["limit_esik"]:
        m.notlar.append(f"Kart limit kullanımı %{limit:g}: ek borçlandırma önerilmez; ödeme planı görüşmesi düşünülebilir")
    f = m.firsatlar.append
    if not evet(m.kart) and not borclandirma_yok and yas >= 18 and gelir >= a["min_gelir_kart"]:
        f(("kart", p["kart"] + (p["maas_bonus"] if maas else 0), f"Kartı yok; aylık gelir {gelir:,.0f} TL".replace(",", ".") + (", maaş müşterisi" if maas else "")))
    if not evet(m.bes) and a["bes_yas"][0] <= yas <= a["bes_yas"][1] and gelir >= a["min_gelir_bes"]:
        f(("bes", p["bes"] + (p["maas_bonus"] if maas else 0), f"BES'i yok; {yas} yaş, düzenli gelir"))
    vadesiz, vadeli, fon = sayi(m.vadesiz), sayi(m.vadeli), sayi(m.fon)
    if vadesiz >= a["vadesiz_esik"] and vadeli == 0 and fon == 0:
        f(("mevduat", p["mevduat"] + (p["yuksek_bakiye_bonus"] if vadesiz >= a["vadesiz_yuksek"] else 0),
           f"3 aylık ort. vadesiz bakiye {vadesiz:,.0f} TL; vadeli veya yatırım ürünü yok".replace(",", ".")))
    vt = tarih(m.vade_tarihi)
    if vadeli > 0 and vt and 0 <= (vt - bugun).days <= a["vade_yenileme_gun"]:
        f(("vade", p["vade"], f"{vadeli:,.0f} TL vadeli mevduatın vadesi {vt:%d.%m.%Y} ({(vt - bugun).days} gün)".replace(",", ".")))
    if evet(m.tasit) and not evet(m.kasko):
        f(("kasko", p["kasko"], "Taşıt kredisi var, bankada kaskosu yok"))
    if evet(m.konut) and not evet(m.konut_sig):
        f(("konut_sig", p["konut_sig"], "Konut kredisi var, konut sigortası yok (DASK zorunlu sigortadır; ayrıca kontrol edin)"))
    if not evet(m.hayat) and a["hayat_yas"][0] <= yas <= a["hayat_yas"][1] and (evet(m.konut) or evet(m.ihtiyac) or evet(m.tasit)):
        f(("hayat", p["hayat"], "Kredisi var, hayat sigortası yok — krediye bağlı zorunlu ürün olarak sunulamaz; müşteri tercihine bırakılmalı"))
    if not evet(m.mobil):
        f(("mobil", p["mobil"], "Mobil bankacılık kullanmıyor"))
    if maas and sayi(m.talimat) == 0:
        f(("talimat", p["talimat"], "Maaş müşterisi, otomatik ödeme talimatı yok"))
    si = tarih(m.son_islem)
    if si and (bugun - si).days > a["pasif_gun"]:
        f(("reaktivasyon", p["reaktivasyon"], f"Son işlem {si:%d.%m.%Y} ({(bugun - si).days} gün önce)"))
    kb = tarih(m.kredi_bitis)
    if kb and 0 <= (kb - bugun).days <= a["kredi_bitis_gun"] and not borclandirma_yok:
        f(("kredi_bitis", p["kredi_bitis"], f"Kredisi {kb:%d.%m.%Y} tarihinde kapanıyor; ihtiyaç görüşmesi (doğrudan kredi önerisi değil)"))
    m.firsatlar.sort(key=lambda x: -x[1])
    m.kanal = "Şube / mobil / arama / SMS" if evet(m.izin) else "Yalnız şube görüşmesinde (pazarlama/İYS izni yok)"
    if not evet(m.izin):
        m.notlar.append("Ticari elektronik ileti (SMS, e-posta, arama) izni yok")


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
UST = Alignment(vertical="top", wrap_text=True)


def _baslik(ws, basliklar, genislik):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w


def rapor_yaz(cikti: Path, musteriler: list[Musteri], ayar: dict, bugun: date, ilk: int, eksik: list[str]) -> None:
    wb = Workbook()
    f = wb.active
    f.title = "Fırsat Listesi"
    _baslik(f, ["Öncelik", "Müşteri No", "Müşteri", "Ürün / Aksiyon", "Puan", "Gerekçe", "Kanal", "Notlar", "Görüşme Tarihi", "Sonuç"],
            (8, 11, 20, 30, 7, 60, 30, 46, 13, 18))
    satirlar = [(m, u, p, g) for m in musteriler for u, p, g in m.firsatlar[:ilk]]
    for i, (m, u, p, g) in enumerate(sorted(satirlar, key=lambda x: (-x[2], -x[0].toplam)), 1):
        f.append([i, m.no, m.ad, URUNLER[u], p, g, m.kanal, "\n".join(m.notlar), None, ""])
        for c in (9, 10):
            f.cell(f.max_row, c).fill = PatternFill("solid", fgColor="FFF4CE")
        if "izni yok" in m.kanal:
            f.cell(f.max_row, 7).fill = PatternFill("solid", fgColor="FDE2E1")
        for h in f[f.max_row]:
            h.alignment = UST
    f.freeze_panes = "D2"
    f.auto_filter.ref = f.dimensions

    o = wb.create_sheet("Müşteri Özeti")
    _baslik(o, ["Müşteri No", "Müşteri", "Ürün Sayısı", "Fırsat Puanı (ilk 3)", "Önerilen Aksiyonlar", "Kanal", "Notlar"], (11, 20, 11, 16, 60, 30, 50))
    for m in sorted(musteriler, key=lambda m: -m.toplam):
        o.append([m.no, m.ad, urun_sayisi(m), m.toplam, "\n".join(f"{URUNLER[u]} ({p})" for u, p, _ in m.firsatlar[:ilk]), m.kanal, "\n".join(m.notlar)])
        for h in o[o.max_row]:
            h.alignment = UST

    u = wb.create_sheet("Ürün Penetrasyonu")
    _baslik(u, ["Ürün", "Sahip Müşteri", "Oran", "Fırsat Sayısı"], (26, 14, 10, 14))
    n = len(musteriler) or 1
    firsat = Counter(x[0] for m in musteriler for x in m.firsatlar)
    ters = {"Kredi kartı": "kart", "BES": "bes", "Vadeli mevduat": "mevduat", "Hayat sigortası": "hayat", "Konut sigortası": "konut_sig",
            "Kasko": "kasko", "Mobil bankacılık": "mobil", "Otomatik ödeme talimatı": "talimat"}
    for ad, a in URUN_SAHIPLIGI:
        sahip = sum(1 for m in musteriler if evet(m.ham.get(a)))
        u.append([ad, sahip, sahip / n, firsat.get(ters.get(ad, ""), None)])
        u.cell(u.max_row, 3).number_format = "0%"
    u.append([])
    u.append(["Ortalama ürün sayısı", round(sum(urun_sayisi(m) for m in musteriler) / n, 2)])
    g = BarChart()
    g.type, g.title, g.height, g.width = "bar", "Ürün sahipliği", 9, 16
    g.add_data(Reference(u, min_col=3, min_row=1, max_row=1 + len(URUN_SAHIPLIGI)), titles_from_data=True)
    g.set_categories(Reference(u, min_col=1, min_row=2, max_row=1 + len(URUN_SAHIPLIGI)))
    u.add_chart(g, "F2")

    k = wb.create_sheet("Kurallar")
    _baslik(k, ["Ayar", "Değer"], (30, 60))
    for a, d in ayar.items():
        k.append([a, json.dumps(d, ensure_ascii=False) if isinstance(d, (dict, list)) else d])
    k.append(["Rapor tarihi", f"{bugun:%d.%m.%Y}"])
    if eksik:
        k.append(["Girdide olmayan sütunlar", ", ".join(eksik) + " (ilgili kurallar uygulanamadı)"])
    k.append([])
    k.append(["Not", "Kurallar ve puanlar örnektir; bankanızın ürün politikası, müşteri uygunluk ve yerindelik (suitability) "
                     "kuralları ile ticari elektronik ileti (İYS) izinleri esastır. Yatırım ürünlerinde müşterinin risk profili "
                     "(uygunluk testi) ayrıca değerlendirilmelidir."])
    k.cell(k.max_row, 2).alignment = UST
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(girdi: Path, cikti: Path, ayar_yolu: Path | None = None, bugun: date | None = None, ilk: int = 3, izinsiz_haric: bool = False) -> dict:
    ayar = json.loads(json.dumps(VARSAYILAN))
    if ayar_yolu:
        ek = json.loads(ayar_yolu.read_text(encoding="utf-8"))
        ayar["puan"].update(ek.pop("puan", {}))
        ayar.update(ek)
    bugun = bugun or date.today()
    musteriler, eksik = portfoy_oku(girdi)
    for m in musteriler:
        degerlendir(m, ayar, bugun)
    if izinsiz_haric:
        musteriler = [m for m in musteriler if evet(m.izin)]
    rapor_yaz(cikti, musteriler, ayar, bugun, ilk, eksik)
    return {"musteriler": musteriler, "eksik": eksik, "ayar": ayar}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Portföydeki bireysel müşteriler için çapraz satış ve elde tutma fırsatlarını önceliklendirir.")
    p.add_argument("--girdi", type=Path, default=ORNEK, help="Portföy listesi (.xlsx/.csv), müşteri başına bir satır")
    p.add_argument("--ayarlar", type=Path, help="Eşik ve puanlar (JSON) — varsayılanların üzerine yazar")
    p.add_argument("--tarih", help="Rapor tarihi GG.AA.YYYY (varsayılan: bugün; örnek veride 08.10.2026)")
    p.add_argument("--ilk", type=int, default=3, help="Müşteri başına listelenecek en fazla fırsat (varsayılan 3)")
    p.add_argument("--izinsiz-haric", action="store_true", help="Pazarlama (İYS) izni olmayan müşterileri listeden çıkar")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "portfoy_firsat_listesi.xlsx")
    a = p.parse_args(argv)
    bugun = tarih(a.tarih) if a.tarih else (date(2026, 10, 8) if a.girdi == ORNEK else None)
    s = calistir(a.girdi, a.cikti, a.ayarlar, bugun, a.ilk, a.izinsiz_haric)
    ms = s["musteriler"]
    print(f"[OK] {len(ms)} müşteri · {sum(len(m.firsatlar[:a.ilk]) for m in ms)} fırsat · ort. ürün sayısı "
          + f"{sum(urun_sayisi(m) for m in ms) / max(len(ms), 1):.1f}".replace(".", ","))
    for m in sorted(ms, key=lambda m: -m.toplam)[:5]:
        print(f"     {m.no} {m.ad}: " + ", ".join(f"{URUNLER[u]} ({p})" for u, p, _ in m.firsatlar[:a.ilk]))
    if s["eksik"]:
        print(f"[i] Girdide olmayan sütunlar (ilgili kurallar uygulanmadı): {', '.join(s['eksik'])}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
