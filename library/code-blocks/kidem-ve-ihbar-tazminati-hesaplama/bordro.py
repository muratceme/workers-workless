"""
bordro.py — Workers / Workless Türkiye bordro çekirdeği

Kaynağı: library/_ortak/bordro.py (paket klasörlerindeki kopyalar CI tarafından aynı tutulur).
Parametreler: tr_parametreler.json (yıl bazında, kaynaklı).

Hesap kuralları:
  - SGK ve işsizlik işçi payları SGK matrahı (brüt, tavanla sınırlı) üzerinden.
  - Gelir vergisi matrahı = brüt − SGK işçi − işsizlik işçi; vergi kümülatif matrah üzerinden
    ücret tarifesiyle hesaplanır (GVK md. 103).
  - Asgari ücret gelir vergisi istisnası (GVK md. 23/18): o takvim ayına kadar biriken asgari ücret
    matrahı üzerinden hesaplanan verginin bu aya düşen kısmı; eksik günde günlük asgari ücret esas alınır.
  - Damga vergisi brüt üzerinden; asgari ücrete isabet eden kısmı istisnadır (DVK md. 9).
  - Tüm ara tutarlar kuruşa yuvarlanır (yarım yukarı).
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

PARAM_DOSYASI = Path(__file__).with_name("tr_parametreler.json")
KURUS = Decimal("0.01")


def para(x) -> Decimal:
    return Decimal(str(x)).quantize(KURUS, ROUND_HALF_UP)


class Parametreler:
    def __init__(self, yil: int | str, dosya: Path = PARAM_DOSYASI):
        veri = json.loads(Path(dosya).read_text(encoding="utf-8"))["yillar"]
        if str(yil) not in veri:
            raise ValueError(f"{yil} yılı parametreleri yok. Mevcut yıllar: {', '.join(sorted(veri))}")
        self.yil = int(yil)
        self.ham = veri[str(yil)]
        D = lambda k: Decimal(self.ham[k])  # noqa: E731
        self.asgari = D("asgari_ucret_brut") if "asgari_ucret_brut" in self.ham else None
        if self.asgari is not None:
            self.sgk_isci, self.issizlik_isci = D("sgk_isci_orani"), D("issizlik_isci_orani")
            self.sgk_isveren, self.issizlik_isveren = D("sgk_isveren_orani"), D("issizlik_isveren_orani")
            self.tesvik = {k: Decimal(v) for k, v in self.ham["sgk_isveren_tesvik_puani"].items()}
            self.taban, self.tavan = D("sgk_taban"), D("sgk_tavan")
        self.tarife = [(Decimal(s) if s is not None else None, Decimal(o))
                       for s, o in self.ham.get("gelir_vergisi_ucret_tarifesi", [])]
        self.damga = D("damga_vergisi_orani")

    def gv_hesapla(self, kumulatif_matrah: Decimal) -> Decimal:
        """Kümülatif matraha ücret tarifesiyle hesaplanan toplam gelir vergisi."""
        if not self.tarife:
            raise ValueError(f"{self.yil} yılı gelir vergisi tarifesi tr_parametreler.json'da yok.")
        vergi, alt = Decimal(0), Decimal(0)
        for ust, oran in self.tarife:
            if ust is None or kumulatif_matrah <= ust:
                return vergi + (kumulatif_matrah - alt) * oran
            vergi += (ust - alt) * oran
            alt = ust
        return vergi

    def dilim_orani(self, kumulatif_matrah: Decimal) -> Decimal:
        for ust, oran in self.tarife:
            if ust is None or kumulatif_matrah <= ust:
                return oran
        return self.tarife[-1][1]

    def kidem_tavani(self, tarih) -> Decimal:
        """Fesih tarihinde yürürlükte olan kıdem tazminatı tavanı (yıllar arası arar)."""
        veri = json.loads(PARAM_DOSYASI.read_text(encoding="utf-8"))["yillar"]
        t = str(tarih)
        for yil in veri.values():
            for donem in yil.get("kidem_tavani", []):
                if donem["baslangic"] <= t <= donem["bitis"]:
                    return Decimal(donem["tutar"])
        raise ValueError(f"{t} tarihi için kıdem tazminatı tavanı tr_parametreler.json'da yok.")

    def au_matrah_aylik(self) -> Decimal:
        return para(self.asgari - para(self.asgari * self.sgk_isci) - para(self.asgari * self.issizlik_isci))


@dataclass
class AyBordrosu:
    ay: int
    gun: int
    brut: Decimal
    sgk_matrahi: Decimal
    sgk_isci: Decimal
    issizlik_isci: Decimal
    gv_matrahi: Decimal
    kumulatif_gv_matrahi: Decimal
    vergi_dilimi: Decimal
    hesaplanan_gv: Decimal
    gv_istisnasi: Decimal
    odenecek_gv: Decimal
    hesaplanan_dv: Decimal
    dv_istisnasi: Decimal
    odenecek_dv: Decimal
    net: Decimal
    sgk_isveren: Decimal
    issizlik_isveren: Decimal
    isveren_maliyeti: Decimal
    uyarilar: str = ""

    def sozluk(self) -> dict:
        return asdict(self)


def ay_hesapla(p: Parametreler, ay: int, brut, onceki_kumulatif=Decimal(0), gun: int = 30, tesvik: str = "yok") -> AyBordrosu:
    """Bir ayın bordrosu. onceki_kumulatif: bu işverende yıl başından önceki ayların toplam GV matrahı."""
    if not 1 <= ay <= 12:
        raise ValueError("Ay 1-12 arasında olmalı")
    if not 1 <= gun <= 30:
        raise ValueError("Gün 1-30 arasında olmalı (SGK ayı 30 gün kabul edilir)")
    if tesvik not in p.tesvik:
        raise ValueError(f"Teşvik türü {', '.join(p.tesvik)} olmalı")
    brut, onceki_kumulatif = para(brut), para(onceki_kumulatif)
    oran_gun = Decimal(gun) / 30
    uyarilar = []

    tavan_gun = para(p.tavan * oran_gun)
    sgk_matrahi = min(brut, tavan_gun)
    if brut > tavan_gun:
        uyarilar.append("SGK tavanı aşıldı; primler tavandan hesaplandı")
    if brut < para(p.taban * oran_gun):
        uyarilar.append("Brüt, SGK prime esas kazanç alt sınırının altında")
    sgk_isci = para(sgk_matrahi * p.sgk_isci)
    issizlik_isci = para(sgk_matrahi * p.issizlik_isci)

    gv_matrahi = brut - sgk_isci - issizlik_isci
    kumulatif = onceki_kumulatif + gv_matrahi
    hesaplanan_gv = para(p.gv_hesapla(kumulatif) - p.gv_hesapla(onceki_kumulatif))

    # Asgari ücret istisnası takvim ayına göre: önceki ayların tam asgari ücret matrahı + bu ayın gün oranı
    au_aylik = p.au_matrah_aylik()
    au_onceki = au_aylik * (ay - 1)
    au_bu_ay = para(au_aylik * oran_gun)
    gv_istisnasi = min(hesaplanan_gv, para(p.gv_hesapla(au_onceki + au_bu_ay) - p.gv_hesapla(au_onceki)))
    odenecek_gv = hesaplanan_gv - gv_istisnasi

    hesaplanan_dv = para(brut * p.damga)
    dv_istisnasi = min(hesaplanan_dv, para(para(p.asgari * oran_gun) * p.damga))
    odenecek_dv = hesaplanan_dv - dv_istisnasi

    net = brut - sgk_isci - issizlik_isci - odenecek_gv - odenecek_dv
    sgk_isveren = para(sgk_matrahi * (p.sgk_isveren - p.tesvik[tesvik]))
    issizlik_isveren = para(sgk_matrahi * p.issizlik_isveren)
    return AyBordrosu(
        ay=ay, gun=gun, brut=brut, sgk_matrahi=sgk_matrahi, sgk_isci=sgk_isci, issizlik_isci=issizlik_isci,
        gv_matrahi=gv_matrahi, kumulatif_gv_matrahi=kumulatif, vergi_dilimi=p.dilim_orani(kumulatif) * 100,
        hesaplanan_gv=hesaplanan_gv, gv_istisnasi=gv_istisnasi, odenecek_gv=odenecek_gv,
        hesaplanan_dv=hesaplanan_dv, dv_istisnasi=dv_istisnasi, odenecek_dv=odenecek_dv, net=net,
        sgk_isveren=sgk_isveren, issizlik_isveren=issizlik_isveren,
        isveren_maliyeti=brut + sgk_isveren + issizlik_isveren, uyarilar="; ".join(uyarilar),
    )


def netten_brute(p: Parametreler, ay: int, net, onceki_kumulatif=Decimal(0), gun: int = 30, tesvik: str = "yok") -> AyBordrosu:
    """Hedef nete ulaşan en küçük brütü kuruş hassasiyetinde bulur (net, brütün azalmayan fonksiyonudur)."""
    hedef = para(net)
    alt, ust = hedef, hedef * 3 + 1000
    while ay_hesapla(p, ay, ust, onceki_kumulatif, gun, tesvik).net < hedef:
        ust *= 2
    alt, ust = int(alt * 100), int(ust * 100)          # kuruş cinsinden ikili arama
    while alt < ust:
        orta = (alt + ust) // 2
        if ay_hesapla(p, ay, Decimal(orta) / 100, onceki_kumulatif, gun, tesvik).net < hedef:
            alt = orta + 1
        else:
            ust = orta
    return ay_hesapla(p, ay, Decimal(alt) / 100, onceki_kumulatif, gun, tesvik)


def yil_hesapla(p: Parametreler, aylar: list[tuple[int, Decimal, int]], tesvik: str = "yok",
                baslangic_kumulatif=Decimal(0), net_hedefli: bool = False) -> list[AyBordrosu]:
    """aylar: [(ay, brüt veya net, gün)]. Kümülatif matrahı aydan aya taşır."""
    sonuc, kum = [], para(baslangic_kumulatif)
    for ay, tutar, gun in aylar:
        b = (netten_brute if net_hedefli else ay_hesapla)(p, ay, tutar, kum, gun, tesvik)
        sonuc.append(b)
        kum = b.kumulatif_gv_matrahi
    return sonuc
