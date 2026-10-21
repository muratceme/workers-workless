"""
Yapay Zekâ API Maliyet Raporu — Gerisi Bende kod bloğu (Sizden Gelenler)

Uygulamanın kendi yapay zekâ API çağrı günlüğünden (CSV / Excel) maliyet, kullanım limiti ve kullanıcı başına maliyet.
  Çağrı maliyeti = günlükteki "maliyet" sütunu varsa o; yoksa fiyat dosyasından:
                   (girdi × girdi fiyatı + çıktı × çıktı fiyatı + önbellek okuma × okuma fiyatı + önbellek yazma × yazma fiyatı) ÷ birim
                   --girdi-onbellek-dahil: girdi token önbellekten okunanları da içeriyorsa (bazı sağlayıcılarda böyle) önce düşülür.
  Model eşleştirme = fiyat dosyasında birebir ad; yoksa adın başıyla eşleşen en uzun anahtar (tarihli sürüm adları için).
                   Fiyatı bulunamayan modelin maliyeti 0 sayılır ve "Yüksek" uyarı verilir.
  Kullanıcı başına = maliyet ÷ o dönemde yapay zekâyı kullanan farklı kullanıcı.
  Ay sonu tahmini = ayın verideki son gününe kadarki harcama ÷ geçen gün × ayın gün sayısı.
  Limitler (isteğe bağlı) = aylık / günlük bütçe, kullanıcı başına günlük istek ve günlük maliyet; aşımlar listelenir.
  Sıçrama = son gün maliyeti ≥ önceki 7 günün ortalamasının 2 katı.
İnternete bağlanmaz; sağlayıcı hesabınıza erişmez.
"""
from __future__ import annotations

import argparse
import calendar
import json
import sys
from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

import tablo as T

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
PAKET = "yapay-zeka-api-maliyet-raporu"
SIFIR = Decimal(0)

CAGRI = {"zaman": ("Zaman", "Tarih", "timestamp", "created_at", "date", "time"), "model": ("Model",),
         "girdi": ("Girdi Token", "input_tokens", "prompt_tokens", "input tokens"), "cikti": ("Çıktı Token", "output_tokens", "completion_tokens", "output tokens"),
         "okuma": ("Önbellek Okuma Token", "cache_read_input_tokens", "cached_tokens", "cache read tokens"),
         "yazma": ("Önbellek Yazma Token", "cache_creation_input_tokens", "cache write tokens"),
         "kullanici": ("Kullanıcı", "user_id", "user", "kullanici id"), "ozellik": ("Özellik", "feature", "endpoint", "islem"),
         "maliyet": ("Maliyet", "cost", "cost_usd"), "durum": ("Durum", "status")}
HATA = {"hata", "error", "failed", "basarisiz", "fail", "timeout"}


def fiyat_oku(yol: Path) -> dict:
    v = json.loads(yol.read_text(encoding="utf-8"))
    modeller = {k: {a: Decimal(str(x)) for a, x in d.items() if a in ("girdi", "cikti", "onbellek_okuma", "onbellek_yazma")}
                for k, d in v.get("modeller", {}).items()}
    return {"para": v.get("para_birimi", "USD"), "birim": Decimal(str(v.get("birim", 1000000))), "modeller": modeller}


def fiyat_bul(fiyat: dict, model: str) -> tuple[str, dict] | tuple[None, None]:
    if model in fiyat["modeller"]:
        return model, fiyat["modeller"][model]
    adaylar = [k for k in fiyat["modeller"] if model.startswith(k)]
    if adaylar:
        k = max(adaylar, key=len)
        return k, fiyat["modeller"][k]
    return None, None


def oku(yollar, fiyat: dict | None, onbellek_dahil: bool, uyarilar) -> list[dict]:
    cagrilar, fiyatsiz = [], defaultdict(int)
    for r in T.hepsini_oku(yollar, CAGRI, ("zaman", "model")):
        z = T.zaman(r["zaman"])
        if z is None:
            uyarilar.append(T.uyari("Orta", "Okunamayan çağrı", f"{r['_dosya']} satır {r['_satir']}", "Zaman okunamadı; atlandı."))
            continue
        n = {a: int(T.sayi(r[a], ".") or 0) for a in ("girdi", "cikti", "okuma", "yazma")}
        model = T.metin(r["model"])
        maliyet, kaynak = T.sayi(r["maliyet"], "."), "günlük"
        if maliyet is None:
            kaynak = "fiyat"
            anahtar, f = fiyat_bul(fiyat, model) if fiyat else (None, None)
            if f is None:
                fiyatsiz[model] += 1
                maliyet = SIFIR
            else:
                girdi = max(0, n["girdi"] - n["okuma"]) if onbellek_dahil else n["girdi"]
                maliyet = (girdi * f.get("girdi", SIFIR) + n["cikti"] * f.get("cikti", SIFIR) + n["okuma"] * f.get("onbellek_okuma", SIFIR)
                           + n["yazma"] * f.get("onbellek_yazma", SIFIR)) / fiyat["birim"]
        cagrilar.append({"zaman": z, "gun": z.date(), "model": model, **n, "kullanici": T.metin(r["kullanici"]) or "—", "ozellik": T.metin(r["ozellik"]) or "—",
                         "maliyet": maliyet, "kaynak": kaynak, "hata": T.katla(r["durum"]) in HATA})
    for model, adet in sorted(fiyatsiz.items()):
        uyarilar.append(T.uyari("Yüksek", "Fiyatı tanımsız model", model or "(boş)",
                                f"{adet} çağrının maliyeti 0 sayıldı. Fiyat dosyasına bu modelin fiyatını ekleyin (birim token başına)."))
    return cagrilar


def analiz(cagrilar: list[dict], limit: dict | None = None, kur: Decimal | None = None) -> dict:
    limit, uyarilar = limit or {}, []
    gunluk = defaultdict(lambda: {"istek": 0, "hata": 0, "girdi": 0, "cikti": 0, "onbellek": 0, "maliyet": SIFIR, "kullanicilar": set()})
    model, ozellik = defaultdict(lambda: {"istek": 0, "girdi": 0, "cikti": 0, "maliyet": SIFIR}), defaultdict(lambda: {"istek": 0, "maliyet": SIFIR})
    kul_gun = defaultdict(lambda: {"istek": 0, "maliyet": SIFIR})
    for c in cagrilar:
        g = gunluk[c["gun"]]
        g["istek"] += 1
        g["hata"] += c["hata"]
        g["girdi"] += c["girdi"]
        g["cikti"] += c["cikti"]
        g["onbellek"] += c["okuma"]
        g["maliyet"] += c["maliyet"]
        g["kullanicilar"].add(c["kullanici"])
        m = model[c["model"]]
        m["istek"] += 1
        m["girdi"] += c["girdi"]
        m["cikti"] += c["cikti"]
        m["maliyet"] += c["maliyet"]
        o = ozellik[c["ozellik"]]
        o["istek"] += 1
        o["maliyet"] += c["maliyet"]
        k = kul_gun[(c["kullanici"], c["gun"])]
        k["istek"] += 1
        k["maliyet"] += c["maliyet"]
    gunler = sorted(gunluk)
    son = gunler[-1]
    toplam = sum((g["maliyet"] for g in gunluk.values()), SIFIR)
    kullanicilar = defaultdict(lambda: {"istek": 0, "maliyet": SIFIR, "gun": 0, "en_yogun": 0, "asim": 0})
    asimlar = []
    for (ku, g), k in sorted(kul_gun.items(), key=lambda x: (x[0][1], x[0][0])):
        x = kullanicilar[ku]
        x["istek"] += k["istek"]
        x["maliyet"] += k["maliyet"]
        x["gun"] += 1
        x["en_yogun"] = max(x["en_yogun"], k["istek"])
        neden = []
        if limit.get("kullanici_gunluk_istek") and k["istek"] > limit["kullanici_gunluk_istek"]:
            neden.append(f"{k['istek']} istek > {limit['kullanici_gunluk_istek']}")
        if limit.get("kullanici_gunluk_maliyet") and k["maliyet"] > Decimal(str(limit["kullanici_gunluk_maliyet"])):
            neden.append(f"maliyet {T.yuvarla(k['maliyet'], '0.0001')} > {limit['kullanici_gunluk_maliyet']}")
        if neden and ku != "—":
            x["asim"] += 1
            asimlar.append({"gun": g, "kullanici": ku, "istek": k["istek"], "maliyet": k["maliyet"], "neden": "; ".join(neden)})
    # ay sonu tahmini: son günün ayı
    ay_gun = [g for g in gunler if (g.year, g.month) == (son.year, son.month)]
    ay_harcama = sum((gunluk[g]["maliyet"] for g in ay_gun), SIFIR)
    ayin_gunu = calendar.monthrange(son.year, son.month)[1]
    tahmin = ay_harcama / son.day * ayin_gunu
    if limit.get("aylik_butce"):
        b = Decimal(str(limit["aylik_butce"]))
        if tahmin > b:
            uyarilar.append(T.uyari("Yüksek", "Aylık bütçe aşılacak", f"{son:%Y-%m}", f"Ay sonu tahmini {T.yuvarla(tahmin)} > bütçe {b} (şu ana kadar {T.yuvarla(ay_harcama)})."))
    if limit.get("gunluk_butce"):
        for g in gunler:
            if gunluk[g]["maliyet"] > Decimal(str(limit["gunluk_butce"])):
                uyarilar.append(T.uyari("Orta", "Günlük bütçe aşıldı", T.tr_tarih(g), f"{T.yuvarla(gunluk[g]['maliyet'])} > {limit['gunluk_butce']}"))
    # veri aralığındaki çağrısız gün sıfır harcamadır; aralığın öncesi hesaba girmez
    onceki = [gunluk[d]["maliyet"] if d in gunluk else SIFIR for d in (son - timedelta(days=i) for i in range(1, 8)) if d >= gunler[0]]
    ort = sum(onceki, SIFIR) / len(onceki) if onceki else SIFIR
    if onceki and gunluk[son]["maliyet"] >= 2 * ort and ort > 0:
        uyarilar.append(T.uyari("Orta", "Maliyet sıçraması", T.tr_tarih(son), f"Son gün {T.yuvarla(gunluk[son]['maliyet'])}, önceki 7 gün ortalaması {T.yuvarla(ort)}."))
    if asimlar:
        uyarilar.append(T.uyari("Orta", "Kullanıcı limiti aşımı", f"{len({a['kullanici'] for a in asimlar})} kullanıcı",
                                f"{len(asimlar)} kullanıcı-gün limit aştı; 'Limit Aşımları' sayfasına bakın."))
    tum_kul = {c["kullanici"] for c in cagrilar if c["kullanici"] != "—"}
    return {"gunluk": dict(gunluk), "model": dict(model), "ozellik": dict(ozellik), "kullanicilar": dict(kullanicilar), "asimlar": asimlar,
            "toplam": toplam, "istek": len(cagrilar), "kullanici_sayisi": len(tum_kul), "kullanici_basina": (toplam / len(tum_kul)) if tum_kul else None,
            "ay": f"{son:%Y-%m}", "ay_harcama": ay_harcama, "tahmin": tahmin, "son": son, "ilk": gunler[0], "kur": kur, "uyarilar": uyarilar, "limit": limit}


def rapor_yaz(cikti: Path, r: dict, para: str, uyarilar) -> None:
    wb = T.yeni_kitap()
    MF = "#,##0.0000"
    ws = T.sayfa(wb, "Özet", ["Gösterge", "Değer", "Açıklama"], (34, 16, 80))
    kurlu = lambda x: (x * r["kur"]) if (x is not None and r["kur"]) else None  # noqa: E731
    for ad, v, b, a in (("Dönem", f"{T.tr_tarih(r['ilk'])} – {T.tr_tarih(r['son'])}", None, ""), (f"Toplam maliyet ({para})", r["toplam"], MF, ""),
                        ("İstek", r["istek"], T.TAMF, ""), ("Yapay zekâyı kullanan kullanıcı", r["kullanici_sayisi"], T.TAMF, "farklı kullanıcı kimliği"),
                        (f"Kullanıcı başına maliyet ({para})", r["kullanici_basina"], MF, "toplam maliyet ÷ kullanan kullanıcı"),
                        (f"{r['ay']} harcaması ({para})", r["ay_harcama"], MF, ""), (f"{r['ay']} ay sonu tahmini ({para})", r["tahmin"], MF, "şu ana kadarki günlük ortalama × ayın gün sayısı"),
                        ("Aylık bütçe", r["limit"].get("aylik_butce"), MF, "limit dosyasından"),
                        ("Kur", r["kur"], "0.0000", f"1 {para} karşılığı; verilirse TL sütunları dolar"), ("Toplam maliyet (TL)", kurlu(r["toplam"]), T.PF, "")):
        T.satir(ws, [ad, v, a], bicim={2: b} if b else None)
    g = T.sayfa(wb, "Günlük", ["Tarih", "İstek", "Hatalı", "Girdi Token", "Çıktı Token", "Önbellek Okuma", f"Maliyet ({para})", "Kullanıcı", "Kullanıcı Başına", "Maliyet (TL)"],
                (12, 9, 9, 13, 13, 14, 14, 10, 14, 13))
    gb = Decimal(str(r["limit"]["gunluk_butce"])) if r["limit"].get("gunluk_butce") else None
    for d, k in sorted(r["gunluk"].items()):
        n = len(k["kullanicilar"] - {"—"})
        T.satir(g, [d, k["istek"], k["hata"], k["girdi"], k["cikti"], k["onbellek"], k["maliyet"], n, (k["maliyet"] / n) if n else None, kurlu(k["maliyet"])],
                bicim={1: T.TF, 4: T.TAMF, 5: T.TAMF, 6: T.TAMF, 7: MF, 9: MF, 10: T.PF}, dolgu={7: T.KIRMIZI if gb and k["maliyet"] > gb else None})
    m = T.sayfa(wb, "Modeller", ["Model", "İstek", "Girdi Token", "Çıktı Token", f"Maliyet ({para})", "Pay"], (30, 9, 14, 14, 14, 8))
    for ad, k in sorted(r["model"].items(), key=lambda x: -x[1]["maliyet"]):
        T.satir(m, [ad, k["istek"], k["girdi"], k["cikti"], k["maliyet"], (k["maliyet"] / r["toplam"]) if r["toplam"] else None], bicim={3: T.TAMF, 4: T.TAMF, 5: MF, 6: T.YF})
    o = T.sayfa(wb, "Özellikler", ["Özellik", "İstek", f"Maliyet ({para})", "İstek Başına", "Pay"], (30, 9, 14, 13, 8))
    for ad, k in sorted(r["ozellik"].items(), key=lambda x: -x[1]["maliyet"]):
        T.satir(o, [ad, k["istek"], k["maliyet"], k["maliyet"] / k["istek"], (k["maliyet"] / r["toplam"]) if r["toplam"] else None], bicim={3: MF, 4: MF, 5: T.YF})
    ku = T.sayfa(wb, "Kullanıcılar", ["Kullanıcı", "İstek", f"Maliyet ({para})", "Aktif Gün", "En Yoğun Gün (istek)", "Limit Aşılan Gün"], (24, 9, 14, 10, 18, 15))
    for ad, k in sorted(r["kullanicilar"].items(), key=lambda x: -x[1]["maliyet"])[:500]:
        T.satir(ku, [ad, k["istek"], k["maliyet"], k["gun"], k["en_yogun"], k["asim"]], bicim={3: MF}, dolgu={6: T.KIRMIZI if k["asim"] else None})
    a = T.sayfa(wb, "Limit Aşımları", ["Tarih", "Kullanıcı", "İstek", f"Maliyet ({para})", "Aşılan Limit"], (12, 24, 9, 14, 50))
    for x in r["asimlar"]:
        T.satir(a, [x["gun"], x["kullanici"], x["istek"], x["maliyet"], x["neden"]], bicim={1: T.TF, 4: MF})
    T.uyari_sayfasi(wb, uyarilar)
    T.kaydet(wb, cikti)


def ozet(r: dict, para: str) -> dict:
    return {"para_birimi": para, "kur": r["kur"], "ay": r["ay"], "ay_sonu_tahmini": r["tahmin"], "aylik_butce": r["limit"].get("aylik_butce"),
            "gunluk": {d.isoformat(): {"maliyet": k["maliyet"], "istek": k["istek"], "kullanici": len(k["kullanicilar"] - {"—"})} for d, k in sorted(r["gunluk"].items())}}


def calistir(girdi, cikti: Path, fiyat_yol: Path | None = None, limit_yol: Path | None = None, kur: Decimal | None = None, onbellek_dahil: bool = False) -> dict:
    uyarilar = []
    fiyat = fiyat_oku(fiyat_yol) if fiyat_yol else None
    cagrilar = oku(girdi, fiyat, onbellek_dahil, uyarilar)
    if not cagrilar:
        raise ValueError("Okunacak çağrı yok.")
    limit = json.loads(limit_yol.read_text(encoding="utf-8")) if limit_yol else {}
    r = analiz(cagrilar, limit, kur)
    r["uyarilar"] = uyarilar + r["uyarilar"]
    para = fiyat["para"] if fiyat else "USD"
    rapor_yaz(cikti, r, para, r["uyarilar"])
    T.ozet_yaz(cikti.parent, PAKET, ozet(r, para))
    r["para"] = para
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Yapay zekâ API çağrı günlüğünden maliyet, limit ve kullanıcı başına maliyet raporu çıkarır.")
    p.add_argument("--girdi", nargs="+", type=Path, help="Çağrı günlüğü (.csv / .xlsx) ya da klasörü: Zaman, Model, Girdi Token, Çıktı Token; isteğe bağlı Kullanıcı, Özellik, Maliyet")
    p.add_argument("--fiyatlar", type=Path, help="Model fiyatları (JSON; birim token başına girdi / çıktı / önbellek fiyatı)")
    p.add_argument("--limitler", type=Path, help="İsteğe bağlı limitler (JSON): aylik_butce, gunluk_butce, kullanici_gunluk_istek, kullanici_gunluk_maliyet")
    p.add_argument("--kur", help="İsteğe bağlı: 1 birim (ör. USD) kaç TL; TL sütunları için")
    p.add_argument("--girdi-onbellek-dahil", action="store_true", help="Girdi token sayısı önbellekten okunanları da içeriyorsa (önce düşülür)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "yapay_zeka_maliyet.xlsx")
    a = p.parse_args(argv)
    if not a.girdi:
        a.girdi = [ORNEK / "api_cagrilari.csv"]
        a.fiyatlar = a.fiyatlar or ORNEK / "fiyatlar.json"
        a.limitler = a.limitler or ORNEK / "limitler.json"
    kur = T.sayi(a.kur) if a.kur else None
    if a.kur and not kur:
        print("[X] --kur bir sayı olmalı (ör. 41,20)")
        return 2
    if not a.fiyatlar:
        print("[!] Fiyat dosyası verilmedi: günlükte 'Maliyet' sütunu yoksa maliyetler 0 sayılır.")
    try:
        r = calistir(a.girdi, a.cikti, a.fiyatlar, a.limitler, kur, a.girdi_onbellek_dahil)
    except (ValueError, json.JSONDecodeError) as h:
        print(f"[X] {h}")
        return 1
    p_ = r["para"]
    print(f"[OK] {T.tr_tarih(r['ilk'])}–{T.tr_tarih(r['son'])} · {r['istek']} istek · maliyet {T.yuvarla(r['toplam'])} {p_} · "
          f"{r['kullanici_sayisi']} kullanıcı · kullanıcı başına {T.yuvarla(r['kullanici_basina'], '0.0001') if r['kullanici_basina'] else '—'} {p_}")
    print(f"[OK] {r['ay']} ay sonu tahmini {T.yuvarla(r['tahmin'])} {p_}" + (f" (bütçe {r['limit']['aylik_butce']})" if r["limit"].get("aylik_butce") else ""))
    for u in r["uyarilar"]:
        if u["onem"] in ("Yüksek", "Orta"):
            print(f"[!] {u['tur']}: {u['kim']} · {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
