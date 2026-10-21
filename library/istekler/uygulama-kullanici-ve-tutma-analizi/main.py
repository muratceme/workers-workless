"""
Uygulama Kullanıcı ve Tutma Analizi — Gerisi Bende kod bloğu (Sizden Gelenler)

Firebase (Google Analytics) BigQuery olay dökümünden günlük kullanıcı, yeni kayıt, özellik kullanımı ve tutma oranları.
  Gün            = event_date (uygulamanın saat dilimi); yoksa event_timestamp'in UTC günü.
  Kullanıcı      = user_pseudo_id (uygulama kurulumu); yoksa user_id.
  Aktif kullanıcı (DAU) = o gün en az bir olayı olan kullanıcı. Arka plan olayları (--pasif-olaylar:
                   notification_receive, notification_dismiss, app_remove) aktiflik saymaz.
  Kohort         = kullanıcının ilk açılış günü: first_open olayının günü; yoksa user_first_touch_timestamp; o da yoksa
                   dökümde ilk görüldüğü gün (bu durumda dökümün ilk günleri "yeni" kullanıcıyı şişirir, uyarı verilir).
  DN tutma       = kohorttaki kullanıcılardan kohort gününden tam N gün sonra aktif olanların payı (D1, D7, D30).
                   Kohort + N günü dökümün son gününden sonraysa hesaplanmaz.
  Ortalama tutma = verisi tamamlanmış kohortların kullanıcı sayısıyla ağırlıklı ortalaması.
  WAU / MAU      = son günde biten 7 / 30 gündeki farklı aktif kullanıcı; yapışkanlık = DAU ÷ MAU.
  Özellik        = otomatik toplanan olaylar (--otomatik listesi) dışındaki olaylar; kullanım ve kullanan kullanıcı.
İnternete bağlanmaz.
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter, defaultdict
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

import tablo as T

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
PAKET = "uygulama-kullanici-ve-tutma-analizi"

OLAY = {"gun": ("event_date", "tarih", "gun"), "zaman": ("event_timestamp",), "olay": ("event_name", "olay"),
        "kullanici": ("user_pseudo_id", "user_id", "kullanici", "kullanici id"), "ilk": ("user_first_touch_timestamp",),
        "platform": ("platform",), "surum": ("app_version", "app_info.version", "app_info_version", "surum"),
        "adet": ("olay_sayisi", "event_count", "adet", "count")}
PASIF = ("notification_receive", "notification_dismiss", "app_remove")
OTOMATIK = ("first_open", "first_visit", "session_start", "user_engagement", "screen_view", "app_update", "os_update", "app_remove", "app_clear_data",
            "app_exception", "notification_receive", "notification_open", "notification_dismiss", "notification_foreground", "firebase_campaign",
            "firebase_in_app_message_action", "firebase_in_app_message_dismiss", "firebase_in_app_message_impression", "in_app_purchase",
            "app_store_subscription_convert", "app_store_subscription_renew", "app_store_refund", "app_store_subscription_cancel",
            "dynamic_link_first_open", "dynamic_link_app_open", "dynamic_link_app_update", "ad_click", "ad_exposure", "ad_impression",
            "ad_query", "ad_reward", "adunit_exposure", "error", "page_view")
GUNLER = (1, 3, 7, 14, 30)


def oku(yollar, uyarilar) -> list[dict]:
    satir = []
    for r in T.hepsini_oku(yollar, OLAY, ("olay", "kullanici")):
        g = T.gun(r["gun"]) or T.gun(r["zaman"])
        k, o = T.metin(r["kullanici"]), T.metin(r["olay"])
        if g is None or not k or not o:
            uyarilar.append(T.uyari("Orta", "Okunamayan satır", f"{r['_dosya']} satır {r['_satir']}", "Gün (event_date / event_timestamp), kullanıcı ya da olay adı yok; atlandı."))
            continue
        adet = T.sayi(r["adet"], ".")
        satir.append({"gun": g, "kullanici": k, "olay": o, "ilk": T.gun(r["ilk"]), "platform": T.metin(r["platform"]).upper() or "—",
                      "surum": T.metin(r["surum"]) or "—", "adet": int(adet) if adet else 1})
    return satir


def analiz(satirlar: list[dict], bitis: date | None = None, pasif=PASIF, otomatik=OTOMATIK) -> dict:
    uyarilar = []
    pasif, otomatik = set(pasif), set(otomatik)
    ilk_gun, son_gun = min(s["gun"] for s in satirlar), max(s["gun"] for s in satirlar)
    bitis = bitis or son_gun
    aktif = defaultdict(set)                     # gün -> kullanıcılar
    oturum, olay_say = Counter(), Counter()
    first_open, ilk_dokunus, ilk_gorulme = {}, {}, {}
    ozellik = defaultdict(lambda: defaultdict(lambda: [0, set()]))    # olay -> gün -> [adet, kullanıcılar]
    platform_gun = defaultdict(set)
    for s in satirlar:
        g, k = s["gun"], s["kullanici"]
        olay_say[g] += s["adet"]
        if s["olay"] == "session_start":
            oturum[g] += s["adet"]
        if s["olay"] == "first_open":
            first_open[k] = min(first_open.get(k, g), g)
        if s["ilk"]:
            ilk_dokunus[k] = min(ilk_dokunus.get(k, s["ilk"]), s["ilk"])
        if s["olay"] not in pasif:
            aktif[g].add(k)
            ilk_gorulme[k] = min(ilk_gorulme.get(k, g), g)
            platform_gun[(g, s["platform"], s["surum"])].add(k)
        if s["olay"] not in otomatik and s["olay"] not in pasif:
            x = ozellik[s["olay"]][g]
            x[0] += s["adet"]
            x[1].add(k)
    kohort, kaynak = {}, Counter()
    for k in ilk_gorulme:
        if k in first_open:
            kohort[k], kaynak["first_open"] = first_open[k], kaynak["first_open"] + 1
        elif k in ilk_dokunus:
            kohort[k], kaynak["user_first_touch_timestamp"] = ilk_dokunus[k], kaynak["user_first_touch_timestamp"] + 1
        else:
            kohort[k], kaynak["ilk görülme"] = ilk_gorulme[k], kaynak["ilk görülme"] + 1
    if kaynak["ilk görülme"]:
        uyarilar.append(T.uyari("Orta", "Kohort ilk görülme gününe göre", f"{kaynak['ilk görülme']} kullanıcı",
                                "Bu kullanıcıların first_open olayı ve user_first_touch_timestamp alanı yok; kohortları dökümde ilk görüldükleri gün sayıldı. "
                                "Dökümün ilk günlerindeki 'yeni kullanıcı' ve tutma oranları bu yüzden yanıltıcı olabilir; dökümü daha geriden alın."))
    gunler = [ilk_gun + timedelta(days=i) for i in range((bitis - ilk_gun).days + 1)]
    yeni = Counter(kohort.values())
    gunluk = {g: {"dau": len(aktif[g]), "yeni": yeni.get(g, 0), "oturum": oturum.get(g, 0), "olay": olay_say.get(g, 0)} for g in gunler}
    for g in gunler:
        if not aktif[g]:
            uyarilar.append(T.uyari("Orta", "Veri olmayan gün", T.tr_tarih(g), "Bu günde hiç aktif kullanıcı yok; döküm eksik olabilir."))
    # kohort tutma
    kohortlar = defaultdict(set)
    for k, c in kohort.items():
        kohortlar[c].add(k)
    tutma = {}
    for c in sorted(kohortlar):
        if c > bitis:
            continue
        uyeler = kohortlar[c]
        tutma[c] = {"kullanici": len(uyeler), **{n: (Decimal(len(uyeler & aktif[c + timedelta(days=n)])) / len(uyeler) if c + timedelta(days=n) <= bitis else None)
                                                  for n in GUNLER}}
    ortalama = {}
    for n in (1, 7, 30):
        tam = [(t["kullanici"], t[n]) for t in tutma.values() if t[n] is not None]
        toplam = sum(k for k, _ in tam)
        ortalama[n] = {"oran": (sum(Decimal(k) * o for k, o in tam) / toplam) if toplam else None, "kohort": len(tam), "kullanici": toplam}
    def penceredeki(gun_sayisi):
        return set().union(*(aktif[bitis - timedelta(days=i)] for i in range(gun_sayisi)))
    wau, mau = len(penceredeki(7)), len(penceredeki(30))
    if (bitis - ilk_gun).days + 1 < 30:
        uyarilar.append(T.uyari("Bilgi", "Kısa döküm", f"{(bitis - ilk_gun).days + 1} gün", "MAU ve D30 için en az 30 günlük (D30 için 60) döküm gerekir."))
    son30 = {bitis - timedelta(days=i) for i in range(30)}
    son7, onceki7 = {bitis - timedelta(days=i) for i in range(7)}, {bitis - timedelta(days=i) for i in range(7, 14)}
    oz = []
    for ad, gler in ozellik.items():
        def kullanan(kume):
            return set().union(*(gler[g][1] for g in kume if g in gler)) if any(g in gler for g in kume) else set()
        k30 = kullanan(son30)
        a7, a_once = sum(gler[g][0] for g in son7 if g in gler), sum(gler[g][0] for g in onceki7 if g in gler)
        oz.append({"ad": ad, "adet": sum(gler[g][0] for g in son30 if g in gler), "kullanici": len(k30),
                   "pay": Decimal(len(k30)) / mau if mau else None, "son7": a7, "onceki7": a_once,
                   "degisim": (Decimal(a7 - a_once) / a_once) if a_once else None})
    oz.sort(key=lambda x: (-x["kullanici"], -x["adet"]))
    if not oz:
        uyarilar.append(T.uyari("Bilgi", "Özellik olayı yok", "", "Dökümde otomatik olaylar dışında olay yok; uygulamada özel olay (logEvent) kaydedilmiyor olabilir."))
    platform = defaultdict(lambda: defaultdict(int))
    for (g, p, s), ku in platform_gun.items():
        if g == bitis:
            platform[p][s] += len(ku)
    return {"ilk": ilk_gun, "bitis": bitis, "gunluk": gunluk, "tutma": tutma, "ortalama": ortalama, "wau": wau, "mau": mau,
            "dau": len(aktif[bitis]), "ozellikler": oz, "platform": platform, "kohort_kaynagi": dict(kaynak), "uyarilar": uyarilar}


def rapor_yaz(cikti: Path, r: dict, uyarilar) -> None:
    wb = T.yeni_kitap()
    ws = T.sayfa(wb, "Özet", ["Gösterge", "Değer", "Açıklama"], (32, 14, 90))
    yap = Decimal(r["dau"]) / r["mau"] if r["mau"] else None
    son7 = [r["gunluk"][r["bitis"] - timedelta(days=i)] for i in range(7) if r["bitis"] - timedelta(days=i) in r["gunluk"]]
    for ad, v, b, a in (("Son gün", r["bitis"], T.TF, ""), ("Aktif kullanıcı (DAU)", r["dau"], T.TAMF, "son gün"),
                        ("Haftalık aktif (WAU)", r["wau"], T.TAMF, "son 7 gün, farklı kullanıcı"), ("Aylık aktif (MAU)", r["mau"], T.TAMF, "son 30 gün, farklı kullanıcı"),
                        ("Yapışkanlık (DAU / MAU)", yap, T.YF, ""), ("Yeni kullanıcı · son 7 gün", sum(g["yeni"] for g in son7), T.TAMF, "kohort gününe göre"),
                        *[(f"D{n} tutma (ortalama)", r["ortalama"][n]["oran"], T.YF, f"{r['ortalama'][n]['kohort']} kohort, {r['ortalama'][n]['kullanici']} kullanıcı; verisi tamamlanmış kohortlar")
                          for n in (1, 7, 30)]):
        T.satir(ws, [ad, v, a], bicim={2: b})
    g = T.sayfa(wb, "Günlük", ["Tarih", "Aktif Kullanıcı", "Yeni Kullanıcı", "Oturum", "Olay", "Oturum / Kullanıcı"], (12, 14, 14, 10, 10, 16))
    for d, k in r["gunluk"].items():
        T.satir(g, [d, k["dau"], k["yeni"], k["oturum"], k["olay"], (Decimal(k["oturum"]) / k["dau"]) if k["dau"] else None], bicim={1: T.TF, 6: "0.00"})
    t = T.sayfa(wb, "Tutma", ["Kohort", "Kullanıcı", *[f"D{n}" for n in GUNLER]], (12, 10, 9, 9, 9, 9, 9))
    for c, k in sorted(r["tutma"].items(), reverse=True):
        T.satir(t, [c, k["kullanici"], *[k[n] for n in GUNLER]], bicim={1: T.TF, **{3 + i: T.YF for i in range(len(GUNLER))}})
    o = T.sayfa(wb, "Özellikler", ["Olay (özellik)", "Kullanım · son 30 gün", "Kullanan · son 30 gün", "MAU payı", "Kullanım · son 7 gün", "Önceki 7 gün", "Değişim"],
                (30, 14, 14, 10, 14, 12, 10))
    for x in r["ozellikler"]:
        T.satir(o, [x["ad"], x["adet"], x["kullanici"], x["pay"], x["son7"], x["onceki7"], x["degisim"]], bicim={4: T.YF, 7: T.YF},
                dolgu={7: (T.KIRMIZI if x["degisim"] is not None and x["degisim"] <= Decimal("-0.3") else None)})
    p = T.sayfa(wb, "Platform ve Sürüm", ["Platform", "Sürüm", "Aktif Kullanıcı (son gün)"], (12, 12, 22))
    for pl, sur in sorted(r["platform"].items()):
        for s, n in sorted(sur.items(), key=lambda x: -x[1]):
            T.satir(p, [pl, s, n])
    T.uyari_sayfasi(wb, uyarilar)
    T.kaydet(wb, cikti)


def ozet(r: dict) -> dict:
    return {"gunluk": {d.isoformat(): k for d, k in r["gunluk"].items()},
            "tutma": {f"d{n}": r["ortalama"][n]["oran"] for n in (1, 7, 30)},
            "wau": r["wau"], "mau": r["mau"],
            "ozellikler": [{"ad": x["ad"], "kullanici": x["kullanici"], "adet": x["adet"]} for x in r["ozellikler"][:10]]}


def calistir(girdi, cikti: Path, bitis: date | None = None, pasif=PASIF, otomatik=OTOMATIK) -> dict:
    uyarilar = []
    satir = oku(girdi, uyarilar)
    if not satir:
        raise ValueError("Okunacak olay yok.")
    r = analiz(satir, bitis, pasif, otomatik)
    r["uyarilar"] = uyarilar + r["uyarilar"]
    rapor_yaz(cikti, r, r["uyarilar"])
    T.ozet_yaz(cikti.parent, PAKET, ozet(r))
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Firebase (Google Analytics) olay dökümünden aktif kullanıcı, yeni kullanıcı, özellik kullanımı ve D1/D7/D30 tutma oranı çıkarır.")
    p.add_argument("--girdi", nargs="+", type=Path, help="Olay dökümü (.csv / .xlsx) ya da klasörü (README'deki sorguyla alınır)")
    p.add_argument("--tarih", help="Son gün GG.AA.YYYY (varsayılan: dökümdeki son gün)")
    p.add_argument("--pasif-olaylar", default=",".join(PASIF), help="Aktiflik saymayan arka plan olayları (virgülle)")
    p.add_argument("--otomatik", default=",".join(OTOMATIK), help="Özellik sayılmayan otomatik olaylar (virgülle)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "kullanici_ve_tutma.xlsx")
    a = p.parse_args(argv)
    a.girdi = a.girdi or [ORNEK / "firebase_olay_ozeti.csv"]
    bitis = T.gun(a.tarih) if a.tarih else None
    if a.tarih and bitis is None:
        print("[X] --tarih GG.AA.YYYY olmalı")
        return 2
    liste = lambda s: tuple(x.strip() for x in s.split(",") if x.strip())  # noqa: E731
    try:
        r = calistir(a.girdi, a.cikti, bitis, liste(a.pasif_olaylar), liste(a.otomatik))
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    o = r["ortalama"]
    yuzde = lambda x: f"%{T.yuvarla(x * 100, '0.1')}" if x is not None else "—"  # noqa: E731
    print(f"[OK] {T.tr_tarih(r['ilk'])}–{T.tr_tarih(r['bitis'])} · DAU {r['dau']} · WAU {r['wau']} · MAU {r['mau']}")
    print(f"[OK] Tutma: D1 {yuzde(o[1]['oran'])} · D7 {yuzde(o[7]['oran'])} · D30 {yuzde(o[30]['oran'])}")
    if r["ozellikler"]:
        print("[OK] En çok kullanılan: " + ", ".join(f"{x['ad']} ({x['kullanici']} kişi)" for x in r["ozellikler"][:5]))
    for u in r["uyarilar"]:
        if u["onem"] in ("Yüksek", "Orta"):
            print(f"[!] {u['tur']}: {u['kim']} · {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
