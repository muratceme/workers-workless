"""
Uygulama Yönetim Özeti — Gerisi Bende kod bloğu (Sizden Gelenler)

Diğer Sizden Gelenler paketlerinin yazdığı <paket>.ozet.json dosyalarını birleştirip günlük, haftalık ya da aylık kısa
yönetim raporu çıkarır (Markdown metin + Excel).
  Dönem      = günlük: o gün · haftalık: son 7 gün · aylık: ay başından o güne. Bitiş günü --tarih, verilmezse verideki
               en son gün.
  Önceki     = günlük: bir önceki gün · haftalık: önceki 7 gün · aylık: önceki ayın aynı günleri.
  Gösterge   = toplam (gelir, çökme, maliyet, talep…) ya da günlük ortalama (DAU); mağaza puanı yorum sayısıyla ağırlıklı.
  Kapsam     = dönemin kaç gününde o kaynağın verisi var; eksik gün varsa karşılaştırma yanıltıcı olabilir.
  Dikkat     = önceki döneme göre kötü yönde en az %20 değişim (--esik), P1 çökme, ay sonu yapay zekâ tahmini > bütçe,
               dönem bitişine kadar uzanmayan (eski) veri.
İnternete bağlanmaz; yalnız verdiğiniz klasörlerdeki .ozet.json dosyalarını okur.
"""
from __future__ import annotations

import argparse
import calendar
import json
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import tablo as T

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
PAKET = "uygulama-yonetim-ozeti"
DONEM_AD = {"gunluk": "Günlük", "haftalik": "Haftalık", "aylik": "Aylık"}

# paket: (bölüm adı, [(alan, gösterge, tür, yön)]) · tür: t toplam, o günlük ortalama, a yorum sayısıyla ağırlıklı ortalama
# yön: +1 artış iyi, -1 artış kötü, 0 yorum yok
KAYNAK = {
    "uygulama-abonelik-gelir-raporu": ("Gelir", [("net", "Net gelir", "t", 1), ("brut_play", "Play brüt satış", "t", 1),
                                                 ("satis_adet", "Satış (adet)", "t", 1), ("iade_adet", "İade (adet)", "t", -1),
                                                 ("iade", "İade tutarı", "t", 0), ("komisyon_play", "Play komisyonu", "t", 0),
                                                 ("iptal", "Abonelik iptali", "t", -1)]),
    "uygulama-kullanici-ve-tutma-analizi": ("Kullanıcı", [("dau", "Günlük aktif kullanıcı (ort.)", "o", 1), ("yeni", "Yeni kullanıcı", "t", 1),
                                                          ("oturum", "Oturum", "t", 1)]),
    "uygulama-cokme-onceliklendirme": ("Çökme", [("fatal_olay", "Çökme (FATAL)", "t", -1), ("anr_olay", "ANR", "t", -1),
                                                 ("nonfatal_olay", "Ölümcül olmayan hata", "t", -1)]),
    "yapay-zeka-api-maliyet-raporu": ("Yapay zekâ", [("maliyet", "API maliyeti", "t", -1), ("istek", "API isteği", "t", 0)]),
    "uygulama-destek-epostasi": ("Destek", [("talep", "Destek e-postası", "t", -1), ("p1", "P1 talep", "t", -1), ("onay", "İnsan onayı gereken", "t", 0)]),
    "magaza-yorumu-siniflandirma": ("Mağaza", [("yorum", "Yorum", "t", 0), ("ortalama_puan", "Ortalama puan", "a", 1), ("olumsuz", "1–2 yıldız", "t", -1)]),
}
SIRA = list(KAYNAK)


def oku(yollar) -> tuple[dict, list[dict]]:
    """Klasörlerdeki (alt klasörler dahil) .ozet.json dosyalarını okur; aynı paketten birden çok varsa en yenisini alır."""
    uyarilar, kaynak = [], {}
    for y in [Path(p) for p in yollar]:
        if not y.exists():
            raise ValueError(f"Bulunamadı: {y}")
        for f in (sorted(y.rglob("*.ozet.json")) if y.is_dir() else [y]):
            try:
                d = json.loads(f.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, UnicodeDecodeError):
                uyarilar.append(T.uyari("Orta", "Okunamayan dosya", f.name, "Geçerli bir .ozet.json değil; atlandı."))
                continue
            p = d.get("paket")
            if p == PAKET or p not in KAYNAK or not isinstance(d.get("gunluk"), dict):
                if p != PAKET:
                    uyarilar.append(T.uyari("Bilgi", "Tanınmayan dosya", f.name, "Sizden Gelenler paketlerinden birine ait değil; atlandı."))
                continue
            d["gunluk"] = {date.fromisoformat(g): v for g, v in d["gunluk"].items()}
            d["_dosya"] = str(f)
            if p in kaynak:
                eski, yeni = sorted([kaynak[p], d], key=lambda x: x.get("uretildi", ""))
                uyarilar.append(T.uyari("Bilgi", "Aynı paketten birden çok dosya", KAYNAK[p][0], f"{Path(yeni['_dosya']).name} kullanıldı (daha yeni)."))
                d = yeni
            kaynak[p] = d
    if not kaynak:
        raise ValueError("Hiç .ozet.json bulunamadı. Önce diğer paketleri çalıştırın; çıktı klasörlerini --kaynak ile verin.")
    return kaynak, uyarilar


def araliklar(t: date, donem: str) -> tuple[tuple[date, date], tuple[date, date]]:
    if donem == "gunluk":
        o = t - timedelta(days=1)
        return (t, t), (o, o)
    if donem == "haftalik":
        return (t - timedelta(days=6), t), (t - timedelta(days=13), t - timedelta(days=7))
    bas = t.replace(day=1)
    onceki_son_ay = bas - timedelta(days=1)
    obas = onceki_son_ay.replace(day=1)
    return (bas, t), (obas, obas.replace(day=min(t.day, calendar.monthrange(obas.year, obas.month)[1])))


def gunler(a: date, b: date) -> list[date]:
    return [a + timedelta(days=i) for i in range((b - a).days + 1)]


def deger(gunluk: dict, aralik, alan: str, tur: str):
    satir = [gunluk[g] for g in gunler(*aralik) if g in gunluk and gunluk[g].get(alan) is not None]
    if not satir:
        return None
    if tur == "o":
        return sum(s[alan] for s in satir) / len(satir)
    if tur == "a":
        agirlik = sum(s.get("yorum") or 0 for s in satir)
        return sum(s[alan] * (s.get("yorum") or 0) for s in satir) / agirlik if agirlik else None
    return sum(s[alan] for s in satir)


def kapsam(gunluk: dict, aralik) -> tuple[int, int]:
    g = gunler(*aralik)
    return sum(1 for x in g if x in gunluk), len(g)


def degisim(bu, once):
    return (bu - once) / abs(once) if bu is not None and once not in (None, 0) else None


def analiz(kaynak: dict, donem: str, t: date | None = None, esik: float = 0.20) -> dict:
    t = t or max(max(d["gunluk"]) for d in kaynak.values())
    bu, once = araliklar(t, donem)
    satirlar, dikkat = [], []
    for p in SIRA:
        if p not in kaynak:
            continue
        d, (bolum, gostergeler) = kaynak[p], KAYNAK[p]
        kap, onceki_kap = kapsam(d["gunluk"], bu), kapsam(d["gunluk"], once)
        birim = d.get("para_birimi") or ""
        for alan, ad, tur, yon in gostergeler:
            b, o = deger(d["gunluk"], bu, alan, tur), deger(d["gunluk"], once, alan, tur)
            dg = degisim(b, o)
            satirlar.append({"bolum": bolum, "ad": ad + (f" ({birim})" if birim and alan in ("net", "brut_play", "iade", "komisyon_play", "maliyet") else ""),
                             "bu": b, "once": o, "degisim": dg, "kapsam": kap, "onceki_kapsam": onceki_kap, "yon": yon, "tur": tur})
            if dg is not None and yon and dg * yon <= -esik and onceki_kap[0] == onceki_kap[1] and kap[0] == kap[1]:
                dikkat.append(T.uyari("Orta", f"{bolum}: {ad}", f"%{dg * 100:+.0f}".replace(".", ","),
                                      f"Önceki döneme göre {'arttı' if dg > 0 else 'azaldı'}: {_tr(b, tur)} ← {_tr(o, tur)}"))
        son = max(d["gunluk"])
        if son < t:
            dikkat.append(T.uyari("Orta", f"{bolum}: eski veri", T.tr_tarih(son), f"Veri {T.tr_tarih(son)} tarihinde bitiyor; dönem {T.tr_tarih(t)}'de bitiyor."))
        elif kap[0] < kap[1]:
            dikkat.append(T.uyari("Bilgi", f"{bolum}: eksik gün", f"{kap[0]}/{kap[1]}", "Dönemin bazı günlerinde veri yok; toplamlar eksik olabilir."))
    ek = ekler(kaynak, bu, once)
    for s in ek["satirlar"]:                          # birleşik göstergeler kendi bölümünün sonuna
        i = max((j for j, x in enumerate(satirlar) if x["bolum"] == s["bolum"]), default=len(satirlar) - 1)
        satirlar.insert(i + 1, s)
    dikkat += ek["dikkat"]
    return {"t": t, "donem": donem, "bu": bu, "once": once, "satirlar": satirlar, "dikkat": dikkat, "anlik": ek["anlik"]}


def ekler(kaynak: dict, bu, once) -> dict:
    """Birden çok paketi birleştiren göstergeler ve dosyaların anlık bilgileri (tutma, kritik çökmeler, ay sonu tahmini…)."""
    satirlar, dikkat, anlik = [], [], []
    ai, kul, gel = (kaynak.get(p) for p in ("yapay-zeka-api-maliyet-raporu", "uygulama-kullanici-ve-tutma-analizi", "uygulama-abonelik-gelir-raporu"))
    if ai and kul:
        def kullanici_gun_basina(aralik):
            ortak = [g for g in gunler(*aralik) if g in ai["gunluk"] and g in kul["gunluk"] and kul["gunluk"][g].get("dau")]
            if not ortak:
                return None
            return sum(ai["gunluk"][g]["maliyet"] for g in ortak) / sum(kul["gunluk"][g]["dau"] for g in ortak)
        b, o = kullanici_gun_basina(bu), kullanici_gun_basina(once)
        satirlar.append({"bolum": "Yapay zekâ", "ad": f"Maliyet / aktif kullanıcı-gün ({ai.get('para_birimi', '')})", "bu": b, "once": o,
                         "degisim": degisim(b, o), "kapsam": None, "onceki_kapsam": None, "yon": -1, "tur": "k"})
    if ai and gel:
        carpan = 1 if ai.get("para_birimi") == gel.get("para_birimi") else (ai.get("kur") if gel.get("para_birimi") == "TRY" and ai.get("kur") else None)
        if carpan:
            def pay(aralik):
                m, n = deger(ai["gunluk"], aralik, "maliyet", "t"), deger(gel["gunluk"], aralik, "net", "t")
                return m * carpan / n if m is not None and n else None
            b, o = pay(bu), pay(once)
            satirlar.append({"bolum": "Yapay zekâ", "ad": "API maliyeti / net gelir", "bu": b, "once": o, "degisim": degisim(b, o), "kapsam": None,
                             "onceki_kapsam": None, "yon": -1, "tur": "y"})
        else:
            anlik.append(("Yapay zekâ", "API maliyeti / net gelir", f"hesaplanmadı: maliyet {ai.get('para_birimi')}, gelir {gel.get('para_birimi')}; "
                                                                    "maliyet raporunu --kur ile çalıştırın"))
    if ai and ai.get("ay_sonu_tahmini") is not None:
        b = ai.get("aylik_butce")
        anlik.append(("Yapay zekâ", f"{ai.get('ay', '')} ay sonu maliyet tahmini", f"{_tr(ai['ay_sonu_tahmini'], 't')} {ai.get('para_birimi', '')}"
                      + (f" (bütçe {_tr(b, 't')})" if b else "")))
        if b and ai["ay_sonu_tahmini"] > b:
            dikkat.append(T.uyari("Yüksek", "Yapay zekâ bütçesi", ai.get("ay", ""), f"Ay sonu tahmini {_tr(ai['ay_sonu_tahmini'], 't')} > bütçe {_tr(b, 't')}"))
    if kul:
        tm = kul.get("tutma") or {}
        if tm:
            anlik.append(("Kullanıcı", "Tutma D1 / D7 / D30", " / ".join(_tr(tm.get(k), "y") for k in ("d1", "d7", "d30"))))
        if kul.get("wau") is not None:
            anlik.append(("Kullanıcı", "Haftalık / aylık aktif (son gün)", f"{kul['wau']} / {kul.get('mau', '—')}"))
    ck = kaynak.get("uygulama-cokme-onceliklendirme")
    if ck:
        anlik.append(("Çökme", "En yeni sürüm", ck.get("en_yeni_surum") or "—"))
        for k in ck.get("kritik") or []:
            yeni = " · yeni sürümde ilk kez" if k.get("yeni") else ""
            anlik.append(("Çökme", f"{k.get('oncelik', 'P1')}: {k.get('baslik', '')}", f"{k.get('tur', '')} · {k.get('kurulum', 0)} kurulum · {k.get('olay', 0)} olay{yeni}"))
            if k.get("oncelik", "P1") == "P1":
                dikkat.append(T.uyari("Yüksek", "Çökme P1", k.get("baslik", ""), f"{k.get('kurulum', 0)} kurulum etkilendi{yeni}"))
    yr = kaynak.get("magaza-yorumu-siniflandirma")
    if yr:
        for o in (yr.get("ozellik_talepleri") or [])[:3]:
            anlik.append(("Mağaza", "En çok istenen", f"{o['ad']} ({o['adet']} yorum)"))
    ds = kaynak.get("uygulama-destek-epostasi")
    if ds:
        for k, n in list((ds.get("kategoriler") or {}).items())[:3]:
            anlik.append(("Destek", "Sık konu", f"{k} ({n})"))
    return {"satirlar": satirlar, "dikkat": dikkat, "anlik": anlik}


def _tr(x, tur="t") -> str:
    if x is None:
        return "—"
    if tur == "y":
        return f"%{x * 100:.1f}".replace(".", ",")
    hane = 4 if tur == "k" else 0 if tur == "t" and float(x).is_integer() else 2
    s = f"{x:,.{hane}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def _degisim_metin(dg) -> str:
    return "—" if dg is None else f"%{dg * 100:+.1f}".replace(".", ",")


def rapor_metin(r: dict) -> str:
    (a, b), (oa, ob) = r["bu"], r["once"]
    ar = lambda x, y: T.tr_tarih(x) if x == y else f"{T.tr_tarih(x)}–{T.tr_tarih(y)}"  # noqa: E731
    m = [f"# Uygulama Yönetim Özeti · {DONEM_AD[r['donem']]} · {ar(a, b)}", "", f"Karşılaştırma: {ar(oa, ob)}", ""]
    if r["dikkat"]:
        m += ["## Dikkat", ""]
        for d in sorted(r["dikkat"], key=lambda x: ["Yüksek", "Orta", "Bilgi"].index(x["onem"])):
            m.append(f"- **{d['tur']}** ({d['kim']}): {d['aciklama']}")
        m.append("")
    bolum = None
    for s in r["satirlar"]:
        if s["bolum"] != bolum:
            bolum = s["bolum"]
            m += ["", f"## {bolum}", ""]
        if s["bu"] is None and s["once"] is None:
            continue
        m.append(f"- {s['ad']}: **{_tr(s['bu'], s['tur'])}** (önceki {_tr(s['once'], s['tur'])}; {_degisim_metin(s['degisim'])})")
    if r["anlik"]:
        m += ["", "## Son durum", ""] + [f"- {b}: {ad}: {d}" for b, ad, d in r["anlik"]]
    m += ["", "_Gerisi Bende · Uygulama Yönetim Özeti. Sayılar kaynak paketlerin ürettiği dosyalardan alınmıştır; karar vermeden önce kontrol edin._", ""]
    return "\n".join(m)


def rapor_yaz(cikti: Path, r: dict, kaynak: dict, uyarilar: list[dict]) -> None:
    wb = T.yeni_kitap()
    ws = T.sayfa(wb, "Özet", ["Bölüm", "Gösterge", "Bu Dönem", "Önceki Dönem", "Değişim", "Kapsam (gün)"], (12, 40, 16, 16, 11, 13))
    for s in r["satirlar"]:
        bic = "0.0%" if s["tur"] == "y" else "#,##0.0000" if s["tur"] == "k" else "#,##0.00"
        kotu = s["degisim"] is not None and s["yon"] and s["degisim"] * s["yon"] < 0
        T.satir(ws, [s["bolum"], s["ad"], s["bu"], s["once"], s["degisim"], f"{s['kapsam'][0]}/{s['kapsam'][1]}" if s["kapsam"] else ""],
                bicim={3: bic, 4: bic, 5: "+0.0%;-0.0%;0.0%"}, dolgu={5: T.KIRMIZI if kotu else None})
    T.satir(ws, [])
    for b, ad, d in r["anlik"]:
        T.satir(ws, [b, ad, d])
    ws = T.sayfa(wb, "Dikkat", ["Önem", "Konu", "Ayrıntı", "Açıklama"], (10, 34, 30, 70))
    for d in sorted(r["dikkat"], key=lambda x: ["Yüksek", "Orta", "Bilgi"].index(x["onem"])):
        T.satir(ws, [d["onem"], d["tur"], d["kim"], d["aciklama"]], dolgu={1: T.ONEM_RENK.get(d["onem"])}, sar=(4,))
    if ws.max_row == 1:
        T.satir(ws, ["—", "Dikkat gerektiren bir şey yok"])
    alanlar = [(p, alan, f"{KAYNAK[p][0]}: {ad}") for p in SIRA if p in kaynak for alan, ad, _, _ in KAYNAK[p][1]]
    ws = T.sayfa(wb, "Günlük", ["Tarih", "Dönem"] + [x[2] for x in alanlar], (12, 10) + (14,) * len(alanlar), "C2")
    for g in gunler(r["once"][0], r["bu"][1]):
        donem = "bu" if g >= r["bu"][0] else "önceki" if g <= r["once"][1] else ""
        T.satir(ws, [g, donem] + [kaynak[p]["gunluk"].get(g, {}).get(alan) for p, alan, _ in alanlar], bicim={1: T.TF})
    ws = T.sayfa(wb, "Kaynaklar", ["Bölüm", "Paket", "Üretildi", "İlk Gün", "Son Gün", "Dosya"], (12, 36, 20, 12, 12, 60))
    for p in SIRA:
        if p in kaynak:
            d = kaynak[p]
            T.satir(ws, [KAYNAK[p][0], p, d.get("uretildi", ""), min(d["gunluk"]), max(d["gunluk"]), d["_dosya"]], bicim={4: T.TF, 5: T.TF})
    eksik = [KAYNAK[p][0] for p in SIRA if p not in kaynak]
    if eksik:
        uyarilar.append(T.uyari("Bilgi", "Eksik kaynak", ", ".join(eksik), "Bu paketlerin .ozet.json dosyası bulunamadı; rapora girmedi."))
    T.uyari_sayfasi(wb, uyarilar)
    T.kaydet(wb, cikti)


def calistir(kaynak_yollari, cikti: Path, donem: str = "gunluk", t: date | None = None, esik: float = 0.20) -> dict:
    kaynak, uyarilar = oku(kaynak_yollari)
    r = analiz(kaynak, donem, t, esik)
    rapor_yaz(cikti, r, kaynak, uyarilar)
    metin = rapor_metin(r)
    cikti.with_suffix(".md").write_text(metin, encoding="utf-8")
    return r | {"metin": metin, "kaynak": kaynak, "uyarilar": uyarilar}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Sizden Gelenler paketlerinin özet dosyalarından günlük / haftalık / aylık yönetim raporu çıkarır.")
    p.add_argument("--kaynak", nargs="+", type=Path, help="Paketlerin çıktı klasörleri ya da .ozet.json dosyaları (alt klasörler de taranır)")
    p.add_argument("--donem", choices=list(DONEM_AD), default="gunluk")
    p.add_argument("--tarih", help="Dönemin son günü (GG.AA.YYYY); verilmezse verideki en son gün")
    p.add_argument("--esik", default="20", help="Dikkat eşiği, yüzde (varsayılan 20)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "yonetim_ozeti.xlsx")
    a = p.parse_args(argv)
    try:
        t = datetime.strptime(a.tarih, "%d.%m.%Y").date() if a.tarih else None
        esik = float(T.sayi(a.esik)) / 100
    except (ValueError, TypeError):
        print("[X] --tarih GG.AA.YYYY, --esik bir sayı olmalı")
        return 2
    try:
        r = calistir(a.kaynak or [ORNEK], a.cikti, a.donem, t, esik)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    print(r["metin"])
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    print(f"[OK] Metin: {a.cikti.with_suffix('.md').resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
