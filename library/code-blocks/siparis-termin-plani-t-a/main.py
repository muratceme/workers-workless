"""
Sipariş Termin Planı (T&A) — Workers / Workless kod bloğu
Tekstil ve Konfeksiyon › Müşteri Temsilciliği (Merchandising) › Production Merchandiser

Sipariş sevk (ex-factory) tarihinden geriye doğru Time & Action (T&A) planı çıkarır:
  - Aşama şablonu: her aşamanın süresi (iş günü, sabit veya adet ÷ günlük kapasite) ve bittiğinde başlayan
    sonraki aşama. Şablon düzenlenebilir (CSV/Excel).
  - Her aşamanın en geç bitiş ve en geç başlangıç tarihi iş takvimiyle (Pazar ve genel tatiller hariç)
    geriye doğru hesaplanır; ilk aşamaların en geç başlangıcı sipariş (PO) tarihinden önceyse termin sıkışıktır.
  - Gerçekleşen tarihler verilirse: zamanında / geç tamamlanan, geciken ve yaklaşan aşamalar; geciken aşamaların
    sevk tarihine olası etkisi (tahmini gecikme).
İnternete bağlanmaz.

Kullanım:
    python main.py                                         # örnek siparişlerle dener
    python main.py --siparisler siparisler.xlsx --sablon ta_sablonu.csv --bugun 08.10.2026
    python main.py --siparisler siparisler.xlsx --gerceklesen gerceklesen.xlsx --calisma-gunu 5
"""
from __future__ import annotations

import argparse
import math
import re
import sys
from collections import OrderedDict, defaultdict
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import puantaj_cekirdek as pc

BURASI = Path(__file__).resolve().parent


def katla(s) -> str:
    return pc.katla(s)


def sayi(x, varsayilan=None):
    if x in (None, ""):
        return varsayilan
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(".", "").replace(",", ".") if "," in str(x) else str(x).strip()
    try:
        return float(s)
    except ValueError:
        return varsayilan


# ----------------------------------------------------------------------------
# İş takvimi
# ----------------------------------------------------------------------------

class Takvim:
    def __init__(self, calisma_gunu: int = 6, ek_tatil: dict[date, str] | None = None):
        self.calisma_gunu = calisma_gunu            # 6: Pzt–Cmt, 5: Pzt–Cum
        self.ek = ek_tatil or {}
        self._tatil: dict[int, set[date]] = {}

    def tatil_mi(self, d: date) -> bool:
        if d.year not in self._tatil:
            tam, _ = pc.tatiller(d.year, self.ek)
            self._tatil[d.year] = set(tam)
        return d in self._tatil[d.year]

    def is_gunu(self, d: date) -> bool:
        return d.weekday() < self.calisma_gunu and not self.tatil_mi(d)

    def geri(self, d: date, gun: int) -> date:
        """d'den geriye 'gun' iş günü (d dahil değil). gun = 0 → d'nin kendisi veya önceki iş günü."""
        while not self.is_gunu(d):
            d -= timedelta(days=1)
        while gun > 0:
            d -= timedelta(days=1)
            if self.is_gunu(d):
                gun -= 1
        return d

    def ilk_is_gunu(self, d: date) -> date:
        while not self.is_gunu(d):
            d += timedelta(days=1)
        return d

    def ileri(self, d: date, gun: int) -> date:
        """d (iş günü) başlangıçlı 'gun' günlük işin bittiği iş günü; gun ≤ 1 → d."""
        d = self.ilk_is_gunu(d)
        while gun > 1:
            d += timedelta(days=1)
            if self.is_gunu(d):
                gun -= 1
        return d

    def fark(self, a: date, b: date) -> int:
        """b − a, iş günü cinsinden (b > a ise pozitif)."""
        if a == b:
            return 0
        isaret = 1 if b > a else -1
        bas, son = (a, b) if b > a else (b, a)
        n, d = 0, bas
        while d < son:
            d += timedelta(days=1)
            if self.is_gunu(d):
                n += 1
        return isaret * n


# ----------------------------------------------------------------------------
# Okuma
# ----------------------------------------------------------------------------

def sablon_oku(yol: Path) -> "OrderedDict[str, dict]":
    s = pc.tablo_oku(yol)
    b = [katla(x) for x in s[0]]
    bul = lambda *a: next((b.index(katla(x)) for x in a if katla(x) in b), None)  # noqa: E731
    i_a, i_s, i_n = bul("aşama", "kilometre taşı", "milestone"), bul("süre gün", "süre", "süre iş günü"), bul("sonraki aşama", "sonraki")
    i_t, i_k = bul("süre türü", "tür"), bul("kapasite adet gün", "kapasite", "günlük kapasite")
    if i_a is None or i_n is None:
        raise SystemExit(f"Şablonda Aşama ve Sonraki Aşama sütunları gerekli. Başlıklar: {s[0]}")
    al = lambda r, i: r[i] if i is not None and i < len(r) else None  # noqa: E731
    sablon: OrderedDict[str, dict] = OrderedDict()
    for r in s[1:]:
        ad = str(al(r, i_a) or "").strip()
        if not ad:
            continue
        tur = "adet" if katla(al(r, i_t)).startswith("adet") else "sabit"
        sablon[ad] = {"ad": ad, "sure": sayi(al(r, i_s), 0.0), "sonraki": str(al(r, i_n) or "").strip(), "tur": tur,
                      "kapasite": sayi(al(r, i_k))}
    son = [a for a, v in sablon.items() if not v["sonraki"]]
    if len(son) != 1:
        raise SystemExit(f"Şablonda sonraki aşaması boş olan tam bir aşama (sevk) olmalı; bulunan: {son}")
    for a, v in sablon.items():
        if v["sonraki"] and v["sonraki"] not in sablon:
            raise SystemExit(f"Şablon: '{a}' aşamasının sonraki aşaması '{v['sonraki']}' tanımlı değil.")
        if v["tur"] == "adet" and not v["kapasite"]:
            raise SystemExit(f"Şablon: '{a}' adete bağlı ama kapasite (adet/gün) verilmemiş.")
    # döngü kontrolü
    for a in sablon:
        gorulen, x = set(), a
        while x:
            if x in gorulen:
                raise SystemExit(f"Şablonda döngü var: {a}")
            gorulen.add(x)
            x = sablon[x]["sonraki"]
    return sablon


def siparisler_oku(yol: Path) -> list[dict]:
    s = pc.tablo_oku(yol)
    b = [katla(x) for x in s[0]]
    bul = lambda *a: next((b.index(katla(x)) for x in a if katla(x) in b), None)  # noqa: E731
    i = {"no": bul("sipariş no", "po no", "sipariş"), "musteri": bul("müşteri", "alıcı", "buyer"), "model": bul("model", "stil", "artikel"),
         "adet": bul("adet", "miktar", "sipariş adedi", "qty"), "po": bul("po tarihi", "sipariş tarihi", "onay tarihi"),
         "sevk": bul("sevk tarihi", "ex factory", "teslim tarihi", "termin")}
    for k, ad in (("no", "Sipariş No"), ("adet", "Adet"), ("sevk", "Sevk Tarihi")):
        if i[k] is None:
            raise SystemExit(f"Sipariş dosyasında '{ad}' sütunu gerekli. Başlıklar: {s[0]}")
    al = lambda r, k: r[i[k]] if i[k] is not None and i[k] < len(r) else None  # noqa: E731
    sonuc = []
    for r in s[1:]:
        if not al(r, "no"):
            continue
        sevk = pc.tarih(al(r, "sevk"))
        if not sevk:
            raise SystemExit(f"{al(r, 'no')}: sevk tarihi okunamadı ({al(r, 'sevk')}).")
        # Ek sütunlar (ör. "Kapasite Dikim") aşama kapasitesini sipariş bazında ezer
        ozel = {}
        for j, h in enumerate(s[0]):
            m = re.match(r"^kapasite\s+(.+)$", katla(h))
            if m and j < len(r) and sayi(r[j]):
                ozel[m.group(1)] = sayi(r[j])
        sonuc.append({"no": str(al(r, "no")).strip(), "musteri": str(al(r, "musteri") or ""), "model": str(al(r, "model") or ""),
                      "adet": sayi(al(r, "adet"), 0.0), "po": pc.tarih(al(r, "po")), "sevk": sevk, "ozel_kapasite": ozel})
    return sonuc


def gerceklesen_oku(yol: Path | None) -> dict[tuple[str, str], date]:
    if not yol:
        return {}
    s = pc.tablo_oku(yol)
    b = [katla(x) for x in s[0]]
    bul = lambda *a: next((b.index(katla(x)) for x in a if katla(x) in b), None)  # noqa: E731
    i_n, i_a, i_t = bul("sipariş no", "po no", "sipariş"), bul("aşama", "kilometre taşı"), bul("gerçekleşen tarih", "tarih", "tamamlanma")
    if None in (i_n, i_a, i_t):
        raise SystemExit(f"Gerçekleşen dosyasında Sipariş No, Aşama, Gerçekleşen Tarih gerekli. Başlıklar: {s[0]}")
    return {(str(r[i_n]).strip(), str(r[i_a]).strip()): pc.tarih(r[i_t]) for r in s[1:]
            if len(r) > max(i_n, i_a, i_t) and r[i_n] and pc.tarih(r[i_t])}


# ----------------------------------------------------------------------------
# Planlama
# ----------------------------------------------------------------------------

def sure(asama: dict, siparis: dict) -> int:
    if asama["tur"] == "adet":
        kap = next((v for k, v in siparis["ozel_kapasite"].items() if katla(asama["ad"]).startswith(k) or k in katla(asama["ad"])),
                   asama["kapasite"])
        return max(1, math.ceil(siparis["adet"] / kap))
    return int(round(asama["sure"]))


def planla(siparis: dict, sablon, takvim: Takvim) -> "OrderedDict[str, dict]":
    """Her aşamanın en geç bitişi = sonraki aşamanın en geç başlangıcı; en geç başlangıç = bitiş − süre (iş günü)."""
    plan: OrderedDict[str, dict] = OrderedDict()

    def hesap(ad: str) -> dict:
        if ad in plan:
            return plan[ad]
        a = sablon[ad]
        g = sure(a, siparis)
        bitis = siparis["sevk"] if not a["sonraki"] else hesap(a["sonraki"])["baslangic"]
        if a["sonraki"]:
            bitis = takvim.geri(bitis, 1)               # sonraki aşama başlamadan önceki iş günü bitmeli
        baslangic = takvim.geri(bitis, max(0, g - 1)) if g else bitis
        plan[ad] = {"asama": ad, "sure": g, "baslangic": baslangic, "bitis": bitis, "sonraki": a["sonraki"]}
        return plan[ad]

    for ad in sablon:
        hesap(ad)
    # Şablon sırasıyla döndür
    return OrderedDict((ad, plan[ad]) for ad in sablon)


def durum_hesapla(siparis: dict, plan, gerceklesen: dict, bugun: date, takvim: Takvim, uyari_gun: int) -> dict:
    # İleri tahmin: tamamlanan aşama gerçekleşen tarihte biter; açık aşama öncülleri bittikten sonraki iş günü başlar
    # (öncülü olmayan aşama PO tarihinde), süresi kadar sürer; tamamlanmadığı için bugünden önce bitmiş sayılamaz.
    oncul: dict[str, list[str]] = defaultdict(list)
    for ad, p in plan.items():
        if p["sonraki"]:
            oncul[p["sonraki"]].append(ad)
    tahmin: dict[str, date] = {}

    def ileri(ad: str) -> date:
        if ad in tahmin:
            return tahmin[ad]
        p = plan[ad]
        g = gerceklesen.get((siparis["no"], ad))
        if g:
            tahmin[ad] = g
            return g
        if oncul[ad]:
            bas = max(ileri(o) for o in oncul[ad]) + timedelta(days=1)
        else:
            bas = siparis["po"] or bugun
        bas = takvim.ilk_is_gunu(bas)
        bitis = takvim.ileri(bas, p["sure"]) if p["sure"] else bas
        tahmin[ad] = max(bitis, takvim.ilk_is_gunu(bugun))
        return tahmin[ad]

    gecikmeler = []
    for ad, p in plan.items():
        g = gerceklesen.get((siparis["no"], ad))
        p["gerceklesen"] = g
        p["tahmini_bitis"] = ileri(ad)
        p["tahmini_gecikme"] = max(0, takvim.fark(p["bitis"], p["tahmini_bitis"]))
        if g:
            gec = takvim.fark(p["bitis"], g)
            p["durum"] = "Tamamlandı" if gec <= 0 else f"Geç tamamlandı ({gec} iş günü)"
            if gec > 0:
                gecikmeler.append((ad, gec, True))
        elif bugun > p["bitis"]:
            gec = takvim.fark(p["bitis"], bugun)
            p["durum"] = f"GECİKTİ ({gec} iş günü)"
            gecikmeler.append((ad, gec, False))
        elif p["tahmini_gecikme"] > 0:
            p["durum"] = f"Gecikecek ({p['tahmini_gecikme']} iş günü)"
        elif takvim.fark(bugun, p["bitis"]) <= uyari_gun:
            p["durum"] = "Yaklaşıyor"
        else:
            p["durum"] = "Planlandı"
    ilkler = [p for ad, p in plan.items() if not oncul[ad]]
    pay = min(takvim.fark(siparis["po"], p["baslangic"]) for p in ilkler) if siparis["po"] else None
    sevk = next(p for p in plan.values() if not p["sonraki"])
    tahmini = sevk["tahmini_gecikme"]
    return {"pay": pay, "tahmini_gecikme": tahmini, "tahmini_sevk": sevk["tahmini_bitis"], "gecikmeler": gecikmeler,
            "risk": "Yüksek" if tahmini > 0 or (pay is not None and pay < 0) else
                    "Orta" if any(p["durum"] == "Yaklaşıyor" for p in plan.values()) or (pay is not None and pay < 5) else "Düşük"}


def calistir(siparis_yolu: Path, sablon_yolu: Path, cikti: Path, gerceklesen_yolu: Path | None = None,
             bugun: date | None = None, calisma_gunu: int = 6, uyari_gun: int = 3) -> dict:
    sablon = sablon_oku(sablon_yolu)
    siparisler = siparisler_oku(siparis_yolu)
    gerceklesen = gerceklesen_oku(gerceklesen_yolu)
    takvim = Takvim(calisma_gunu)
    bugun = bugun or date.today()
    sonuc = []
    for s in siparisler:
        plan = planla(s, sablon, takvim)
        d = durum_hesapla(s, plan, gerceklesen, bugun, takvim, uyari_gun)
        sonuc.append({"siparis": s, "plan": plan, **d})
    _rapor(sonuc, sablon, bugun, calisma_gunu, cikti)
    return {"sonuc": sonuc, "bugun": bugun}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
DURUM_DOLGU = {"GECİKTİ": "F8C9C6", "Gecikecek": "FDE2E1", "Geç": "FDE2E1", "Yaklaşıyor": "FFF4CE", "Tamamlandı": "E3F5E1",
               "Planlandı": "FFFFFF"}
RISK_DOLGU = {"Yüksek": "F8C9C6", "Orta": "FFF4CE", "Düşük": "E3F5E1"}
TARIH = "DD.MM.YYYY"


def _baslik(ws, satir=1):
    for c in ws[satir]:
        c.fill, c.font = BASLIK_DOLGU, BASLIK_YAZI


def _dolgu(durum: str) -> PatternFill:
    for k, v in DURUM_DOLGU.items():
        if durum.startswith(k):
            return PatternFill("solid", fgColor=v)
    return PatternFill()


def _rapor(sonuc, sablon, bugun, calisma_gunu, cikti):
    wb = Workbook()
    o = wb.active
    o.title = "Sipariş Özeti"
    o.append([f"T&A durumu · {bugun:%d.%m.%Y}"])
    o["A1"].font = Font(bold=True, size=12)
    o.append(["Sipariş No", "Müşteri", "Model", "Adet", "PO Tarihi", "Sevk Tarihi", "Plan Payı (iş günü)", "Geciken Aşama",
              "Tahmini Sevk", "Tahmini Sevk Gecikmesi (iş günü)", "Risk", "Sıradaki Aşama", "Son Tarihi"])
    _baslik(o, 2)
    for x in sorted(sonuc, key=lambda x: ({"Yüksek": 0, "Orta": 1, "Düşük": 2}[x["risk"]], x["siparis"]["sevk"])):
        s = x["siparis"]
        sira = next((p for p in sorted(x["plan"].values(), key=lambda p: p["bitis"]) if not p["gerceklesen"]), None)
        o.append([s["no"], s["musteri"], s["model"], s["adet"], s["po"], s["sevk"], x["pay"],
                  ", ".join(a for a, _, t in x["gecikmeler"] if not t) or None, x["tahmini_sevk"], x["tahmini_gecikme"] or None,
                  x["risk"], sira["asama"] if sira else "-", sira["bitis"] if sira else None])
        for c in (5, 6, 9, 13):
            o.cell(o.max_row, c).number_format = TARIH
        o.cell(o.max_row, 11).fill = PatternFill("solid", fgColor=RISK_DOLGU[x["risk"]])
        if x["pay"] is not None and x["pay"] < 0:
            o.cell(o.max_row, 7).fill = PatternFill("solid", fgColor=RISK_DOLGU["Yüksek"])
    for j, w in enumerate((14, 16, 12, 9, 12, 12, 11, 30, 12, 14, 8, 22, 12), 1):
        o.column_dimensions[get_column_letter(j)].width = w
    o.freeze_panes = "B3"

    m = wb.create_sheet("T&A Matrisi")
    asamalar = list(sablon)
    m.append(["Sipariş No", "Sevk"] + asamalar)
    _baslik(m)
    for x in sonuc:
        m.append([x["siparis"]["no"], x["siparis"]["sevk"]] + [x["plan"][a]["bitis"] for a in asamalar])
        m.cell(m.max_row, 2).number_format = TARIH
        for j, a in enumerate(asamalar, 3):
            c = m.cell(m.max_row, j)
            c.number_format = TARIH
            c.fill = _dolgu(x["plan"][a]["durum"])
    m.column_dimensions["A"].width = 14
    for j in range(2, len(asamalar) + 3):
        m.column_dimensions[get_column_letter(j)].width = 13
    m.freeze_panes = "C2"
    m.append([])
    m.append(["Hücreler aşamanın en geç bitiş tarihidir. Renk: kırmızı gecikti, sarı yaklaşıyor, yeşil tamamlandı."])

    d = wb.create_sheet("Aşama Detayı")
    d.append(["Sipariş No", "Aşama", "Süre (iş günü)", "En Geç Başlangıç", "En Geç Bitiş", "Gerçekleşen", "Tahmini Bitiş", "Durum",
              "Sonraki Aşama"])
    _baslik(d)
    for x in sonuc:
        for p in sorted(x["plan"].values(), key=lambda p: (p["bitis"], p["baslangic"])):
            d.append([x["siparis"]["no"], p["asama"], p["sure"], p["baslangic"], p["bitis"], p["gerceklesen"], p["tahmini_bitis"],
                      p["durum"], p["sonraki"]])
            for c in (4, 5, 6, 7):
                d.cell(d.max_row, c).number_format = TARIH
            d.cell(d.max_row, 8).fill = _dolgu(p["durum"])
    for j, w in enumerate((14, 24, 10, 14, 14, 13, 13, 26, 22), 1):
        d.column_dimensions[get_column_letter(j)].width = w
    d.freeze_panes = "C2"
    d.auto_filter.ref = d.dimensions

    b = wb.create_sheet("Bilgi")
    for s in [["Yöntem", "Sevk tarihinden geriye: her aşamanın en geç bitişi, sonraki aşamanın en geç başlangıcından önceki iş günüdür"],
              ["Süre", "sabit iş günü veya adet ÷ günlük kapasite (yukarı yuvarlanır); siparişe özel kapasite 'Kapasite Dikim' gibi sütunla verilir"],
              ["Takvim", f"haftada {calisma_gunu} iş günü (Pazartesi'den başlayarak); sabit genel tatiller ve dini bayramlar iş günü sayılmaz"],
              ["Plan payı", "ilk aşamaların en geç başlangıcı ile PO tarihi arasındaki iş günü; negatifse termin bu şablonla yetişmez"],
              ["Tahmini bitiş", "tamamlanan aşama gerçekleşen tarihte; açık aşama öncülleri bittikten sonraki iş günü (en erken bugün) "
                                "başlar ve süresi kadar sürer. Sevk aşamasının tahmini bitişi ile son tarihi arasındaki fark tahmini gecikmedir; "
                                "işler hızlandırılmazsa sevk bu kadar kayar"],
              ["Not", "Şablondaki süreler örnektir; kumaş, boyahane ve atölye tedarik sürelerinizle güncelleyin. Kesim, dikim ve "
                      "ütü-paket uygulamada kısmen paralel yürür; bu planlama tampon içeren ihtiyatlı (seri) bir plandır."]]:
        b.append(s)
    b.column_dimensions["A"].width = 16
    b.column_dimensions["B"].width = 130
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ornek = BURASI / "ornek_veri"
    ap = argparse.ArgumentParser(description="Sevk tarihinden geriye doğru T&A (Time & Action) planı çıkarır ve gecikme riskini işaretler.")
    ap.add_argument("--siparisler", type=Path, default=ornek / "siparisler.csv",
                    help="Siparişler (.xlsx/.csv): Sipariş No, Müşteri, Model, Adet, PO Tarihi, Sevk Tarihi [, Kapasite <Aşama>]")
    ap.add_argument("--sablon", type=Path, default=BURASI / "ta_sablonu.csv",
                    help="Aşama şablonu: Aşama, Süre (gün), Sonraki Aşama, Süre Türü (sabit/adet), Kapasite (adet/gün)")
    ap.add_argument("--gerceklesen", type=Path, help="Gerçekleşen tarihler: Sipariş No, Aşama, Gerçekleşen Tarih")
    ap.add_argument("--bugun", help="Durum tarihi (GG.AA.YYYY); varsayılan bugün")
    ap.add_argument("--calisma-gunu", type=int, default=6, choices=[5, 6, 7], help="Haftalık iş günü (varsayılan 6: Pzt–Cmt)")
    ap.add_argument("--uyari-gun", type=int, default=3, help="Son tarihine bu kadar iş günü kalan aşama 'yaklaşıyor' (varsayılan 3)")
    ap.add_argument("--cikti", type=Path, default=Path("cikti") / "ta_plani.xlsx")
    a = ap.parse_args(argv)
    bugun = pc.tarih(a.bugun) if a.bugun else None
    if a.siparisler == ornek / "siparisler.csv":
        a.gerceklesen = a.gerceklesen or ornek / "gerceklesen.csv"
        bugun = bugun or date(2026, 10, 8)
    s = calistir(a.siparisler, a.sablon, a.cikti, a.gerceklesen, bugun, a.calisma_gunu, a.uyari_gun)
    for x in s["sonuc"]:
        sp = x["siparis"]
        print(f"[OK] {sp['no']} ({sp['adet']:g} adet, sevk {sp['sevk']:%d.%m.%Y}): risk {x['risk']}"
              + (f" · plan payı {x['pay']} iş günü" if x["pay"] is not None else "")
              + (f" · tahmini gecikme {x['tahmini_gecikme']} iş günü" if x["tahmini_gecikme"] else ""))
    print(f"[OK] Rapor: {a.cikti.resolve()}")


if __name__ == "__main__":
    sys.exit(main())
