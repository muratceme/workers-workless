"""
Uygulama Çökme Önceliklendirme — Gerisi Bende kod bloğu (Sizden Gelenler)

Firebase Crashlytics'in BigQuery dökümünden (CSV / Excel) sorunları önceliklendirir.
  Sorun (issue)      = issue_id; başlık, dökümdeki başlık sütunlarından en sık görülen değerdir (yoksa issue_id).
  Tür                = error_type: FATAL (çökme), ANR (uygulama yanıt vermiyor), NON_FATAL. Eski dökümlerde is_fatal.
  Etkilenen kurulum  = farklı installation_uuid sayısı (Crashlytics kurulum bazlı sayar; kullanıcıya yaklaşık eşittir).
  Pencere            = --tarih'te biten son --gun gün (varsayılan: dökümdeki son gün, 7 gün).
  En yeni sürüm      = penceredeki en büyük application.display_version (sayısal parçalara göre: 2.10 > 2.9).
  Yeni sorun         = ilk kez en yeni sürümde ve pencere içinde görülen sorun.
  Sıçrama            = son gündeki olay ≥ önceki günlerin günlük ortalamasının 2 katı ve ≥ 5 olay.
  Öncelik (kod kuralı):
    P1  FATAL / ANR ve (etkilenen kurulum ≥ --esik  ya da  yeni sorun ve ≥ 3 kurulum  ya da  sıçrama)
    P2  diğer FATAL / ANR;  NON_FATAL'da sıçrama ya da etkilenen kurulum ≥ 5 × --esik
    P3  diğerleri
  Çökmesiz kullanıcı = 1 − (sürümde çöken kurulum ÷ sürümün aktif kullanıcısı); yalnız --kullanicilar verilirse, yaklaşık.
İnternete bağlanmaz; bildirim göndermez, yapıştırmaya hazır bildirim metni yazar.
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import Counter, defaultdict
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

import tablo as T

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
PAKET = "uygulama-cokme-onceliklendirme"

OLAY = {"zaman": ("event_timestamp", "zaman", "tarih"), "issue": ("issue_id", "sorun"), "tur": ("error_type", "hata turu"), "fatal": ("is_fatal",),
        "kurulum": ("installation_uuid", "kurulum"), "surum": ("display_version", "application.display_version", "application_display_version", "app_version", "surum"),
        "platform": ("platform",), "cihaz": ("device_model", "device.model", "cihaz"), "os": ("os_version", "operating_system.display_version", "isletim sistemi"),
        "baslik": ("title", "issue_title", "baslik", "exceptions.type", "exception_type", "error.title", "blame_symbol", "blame_frame.symbol"),
        "olay_id": ("event_id",)}
KULLANICI = {"surum": ("Sürüm", "display_version", "app_info.version", "Version"), "kullanici": ("Aktif Kullanıcı", "Kullanıcı", "active_users", "Users")}
TUR_SIRA = {"FATAL": 0, "ANR": 1, "NON_FATAL": 2}
TUR_AD = {"FATAL": "Çökme", "ANR": "ANR (yanıt vermiyor)", "NON_FATAL": "Çökmeyen hata"}
P_RENK = {"P1": T.KIRMIZI, "P2": T.SARI, "P3": T.GRI}


def surum_anahtari(s: str):
    return tuple(int(x) if x.isdigit() else -1 for x in re.split(r"[.\-_ ]", s or "")) or (-1,)


def _tur(r) -> str | None:
    t = T.katla(r["tur"]).replace(" ", "_").upper()
    if t in TUR_SIRA:
        return t
    f = T.katla(r["fatal"])
    if f in ("true", "1", "evet"):
        return "FATAL"
    if f in ("false", "0", "hayir"):
        return "NON_FATAL"
    return None


def oku(yollar, uyarilar) -> list[dict]:
    olaylar, gorulen = [], set()
    for r in T.hepsini_oku(yollar, OLAY, ("zaman", "issue", "kurulum")):
        z, tur = T.zaman(r["zaman"]), _tur(r)
        if z is None or tur is None:
            uyarilar.append(T.uyari("Orta", "Okunamayan olay", f"{r['_dosya']} satır {r['_satir']}",
                                    "Zaman ya da tür (error_type / is_fatal) okunamadı; satır atlandı."))
            continue
        oid = T.metin(r["olay_id"])
        if oid:
            if oid in gorulen:
                continue
            gorulen.add(oid)
        olaylar.append({"zaman": z, "gun": z.date(), "issue": T.metin(r["issue"]), "tur": tur, "kurulum": T.metin(r["kurulum"]),
                        "surum": T.metin(r["surum"]) or "?", "platform": T.metin(r["platform"]).upper(), "cihaz": T.metin(r["cihaz"]),
                        "os": T.metin(r["os"]), "baslik": T.metin(r["baslik"])})
    return olaylar


def oku_kullanicilar(yol: Path | None) -> dict:
    if not yol:
        return {}
    k = defaultdict(int)
    for r in T.kayitlar(yol, KULLANICI, ("surum", "kullanici")):
        n = T.sayi(r["kullanici"])
        if T.metin(r["surum"]) and n:
            k[T.metin(r["surum"])] += int(n)
    return dict(k)


def analiz(olaylar: list[dict], bitis: date, gun: int = 7, esik: int = 10, kullanicilar: dict | None = None) -> dict:
    bas = bitis - timedelta(days=gun - 1)
    pencere = [o for o in olaylar if bas <= o["gun"] <= bitis]
    if not pencere:
        raise ValueError(f"{T.tr_tarih(bas)}–{T.tr_tarih(bitis)} arasında olay yok. --tarih ve --gun değerlerini kontrol edin.")
    en_yeni = max({o["surum"] for o in pencere if o["surum"] != "?"} or {"?"}, key=surum_anahtari)
    ilk = {}
    for o in sorted(olaylar, key=lambda o: o["zaman"]):
        ilk.setdefault(o["issue"], o)
    gruplar = defaultdict(list)
    for o in pencere:
        gruplar[o["issue"]].append(o)
    sorunlar = []
    for iid, ol in gruplar.items():
        tur = min((o["tur"] for o in ol), key=TUR_SIRA.get)
        basliklar = Counter(o["baslik"] for o in ol if o["baslik"])
        gunluk = Counter(o["gun"] for o in ol)
        son, onceki = gunluk.get(bitis, 0), [gunluk.get(bitis - timedelta(days=i), 0) for i in range(1, gun)]
        ort = Decimal(sum(onceki)) / len(onceki) if onceki else Decimal(0)
        sicrama = son >= 5 and son >= 2 * ort
        kurulum = len({o["kurulum"] for o in ol})
        f = ilk[iid]
        yeni = f["surum"] == en_yeni and f["gun"] >= bas
        if tur in ("FATAL", "ANR") and (kurulum >= esik or (yeni and kurulum >= 3) or sicrama):
            p = "P1"
        elif tur in ("FATAL", "ANR") or sicrama or kurulum >= 5 * esik:
            p = "P2"
        else:
            p = "P3"
        cokme = tur in ("FATAL", "ANR")
        kurulum_esigi = esik if cokme else 5 * esik
        neden = [x for x, v in ((f"{kurulum} kurulum etkilendi (eşik {kurulum_esigi})", kurulum >= kurulum_esigi), ("en yeni sürümde ortaya çıktı", yeni),
                                (f"son gün {son} olay, önceki günler ortalama {T.yuvarla(ort, '0.1')}", sicrama)) if v]
        if not neden and p == "P2":
            neden = ["çökme / ANR olduğu için en az P2; P1 eşiklerinin altında"]
        sorunlar.append({"issue": iid, "baslik": basliklar.most_common(1)[0][0] if basliklar else iid, "tur": tur, "olay": len(ol), "kurulum": kurulum,
                         "ilk": f["zaman"], "ilk_surum": f["surum"], "son": max(o["zaman"] for o in ol), "surumler": sorted({o["surum"] for o in ol}, key=surum_anahtari),
                         "platform": ", ".join(sorted({o["platform"] for o in ol if o["platform"]})), "son_gun": son, "ort": ort, "sicrama": sicrama, "yeni": yeni,
                         "oncelik": p, "neden": "; ".join(neden) or "eşiklerin altında",
                         "cihaz": ", ".join(c for c, _ in Counter(o["cihaz"] for o in ol if o["cihaz"]).most_common(3))})
    sorunlar.sort(key=lambda s: (s["oncelik"], TUR_SIRA[s["tur"]], -s["kurulum"], -s["olay"]))
    surumler = {}
    for s in sorted({o["surum"] for o in pencere}, key=surum_anahtari, reverse=True):
        ol = [o for o in pencere if o["surum"] == s]
        cokme = {o["kurulum"] for o in ol if o["tur"] in ("FATAL", "ANR")}
        aktif = (kullanicilar or {}).get(s)
        surumler[s] = {"olay": len(ol), "cokme_olay": sum(o["tur"] in ("FATAL", "ANR") for o in ol), "cokme_kurulum": len(cokme),
                       "aktif": aktif, "cokmesiz": (1 - Decimal(len(cokme)) / aktif) if aktif else None}
    gunluk = {}
    for i in range(gun):
        d = bas + timedelta(days=i)
        ol = [o for o in pencere if o["gun"] == d]
        gunluk[d] = {"fatal_olay": sum(o["tur"] == "FATAL" for o in ol), "anr_olay": sum(o["tur"] == "ANR" for o in ol),
                     "nonfatal_olay": sum(o["tur"] == "NON_FATAL" for o in ol), "etkilenen_kurulum": len({o["kurulum"] for o in ol if o["tur"] in ("FATAL", "ANR")})}
    return {"bas": bas, "bitis": bitis, "en_yeni": en_yeni, "sorunlar": sorunlar, "surumler": surumler, "gunluk": gunluk, "esik": esik}


def bildirim_metni(r: dict) -> str:
    kritik = [s for s in r["sorunlar"] if s["oncelik"] == "P1"]
    bas = f"Çökme özeti {T.tr_tarih(r['bas'])}–{T.tr_tarih(r['bitis'])} · en yeni sürüm {r['en_yeni']}"
    g = r["gunluk"][r["bitis"]]
    satirlar = [bas, f"Son gün: {g['fatal_olay']} çökme, {g['anr_olay']} ANR, {g['etkilenen_kurulum']} kurulum etkilendi."]
    if not kritik:
        satirlar.append("P1 sorun yok.")
    for s in kritik[:5]:
        ek = [n for n in s["neden"].split("; ") if "kurulum etkilendi" not in n]   # kurulum sayısı zaten satırda
        satirlar.append(f"• P1 {TUR_AD[s['tur']]}: {s['baslik']} — {s['kurulum']} kurulum, {s['olay']} olay" + (f" ({'; '.join(ek)})" if ek else ""))
    if len(kritik) > 5:
        satirlar.append(f"… ve {len(kritik) - 5} P1 sorun daha (Excel raporunda).")
    return "\n".join(satirlar) + "\n"


def rapor_yaz(cikti: Path, r: dict, uyarilar: list[dict]) -> None:
    wb = T.yeni_kitap()
    bas = ["Öncelik", "Tür", "Sorun", "Etkilenen Kurulum", "Olay", "Son Gün", "Neden", "İlk Görülme", "İlk Sürüm", "Son Görülme", "Sürümler",
           "Platform", "Sık Cihazlar", "issue_id"]
    gen = (8, 18, 46, 10, 8, 8, 50, 16, 9, 16, 20, 12, 30, 34)
    for ad, filtre in (("Kritik", lambda s: s["oncelik"] in ("P1", "P2")), ("Sorunlar", lambda s: True)):
        ws = T.sayfa(wb, ad, bas, gen, "D2")
        for s in filter(filtre, r["sorunlar"]):
            T.satir(ws, [s["oncelik"], TUR_AD[s["tur"]], s["baslik"], s["kurulum"], s["olay"], s["son_gun"], s["neden"], s["ilk"], s["ilk_surum"], s["son"],
                         ", ".join(s["surumler"]), s["platform"], s["cihaz"], s["issue"]],
                    bicim={8: T.ZF, 10: T.ZF}, dolgu={1: P_RENK[s["oncelik"]], 9: T.SARI if s["yeni"] else None}, sar=(3, 7))
    sv = T.sayfa(wb, "Sürümler", ["Sürüm", "Olay", "Çökme + ANR Olayı", "Çöken Kurulum", "Aktif Kullanıcı", "Çökmesiz Kullanıcı (yaklaşık)"], (10, 8, 14, 12, 12, 16))
    for s, k in r["surumler"].items():
        T.satir(sv, [s, k["olay"], k["cokme_olay"], k["cokme_kurulum"], k["aktif"], k["cokmesiz"]], bicim={6: "0.00%"},
                dolgu={1: T.SARI if s == r["en_yeni"] else None})
    g = T.sayfa(wb, "Günlük", ["Tarih", "Çökme", "ANR", "Çökmeyen Hata", "Etkilenen Kurulum"], (12, 9, 9, 13, 14))
    for d, k in r["gunluk"].items():
        T.satir(g, [d, k["fatal_olay"], k["anr_olay"], k["nonfatal_olay"], k["etkilenen_kurulum"]], bicim={1: T.TF})
    k = T.sayfa(wb, "Kurallar", ["Kural", "Açıklama"], (22, 110))
    for x in (("P1", f"Çökme / ANR ve (etkilenen kurulum ≥ {r['esik']} ya da en yeni sürümde ortaya çıkan ve ≥ 3 kurulum ya da sıçrama)"),
              ("P2", f"Diğer çökme / ANR; çökmeyen hatada sıçrama ya da ≥ {5 * r['esik']} kurulum"), ("P3", "Diğerleri"),
              ("Sıçrama", "Son gündeki olay ≥ önceki günlerin ortalamasının 2 katı ve ≥ 5"),
              ("Çökmesiz kullanıcı", "1 − çöken kurulum ÷ sürümün aktif kullanıcısı (kurulum ≈ kullanıcı varsayımıyla, yaklaşık)")):
        T.satir(k, list(x), sar=(2,))
    T.uyari_sayfasi(wb, uyarilar)
    T.kaydet(wb, cikti)


def ozet(r: dict) -> dict:
    return {"en_yeni_surum": r["en_yeni"], "gunluk": {d.isoformat(): v for d, v in r["gunluk"].items()},
            "kritik": [{"baslik": s["baslik"], "tur": s["tur"], "kurulum": s["kurulum"], "olay": s["olay"], "oncelik": s["oncelik"], "yeni": s["yeni"]}
                       for s in r["sorunlar"] if s["oncelik"] == "P1"][:10],
            "sayilar": dict(Counter(s["oncelik"] for s in r["sorunlar"]))}


def calistir(girdi, cikti: Path, bitis: date | None = None, gun: int = 7, esik: int = 10, kullanici_yol: Path | None = None) -> dict:
    uyarilar = []
    olaylar = oku(girdi, uyarilar)
    if not olaylar:
        raise ValueError("Okunacak çökme olayı yok.")
    bitis = bitis or max(o["gun"] for o in olaylar)
    r = analiz(olaylar, bitis, gun, esik, oku_kullanicilar(kullanici_yol))
    if r["en_yeni"] == "?":
        uyarilar.append(T.uyari("Orta", "Sürüm sütunu yok", "", "application.display_version bulunamadı; 'yeni sorun' kuralı çalışmaz."))
    if not any(o["baslik"] for o in olaylar):
        uyarilar.append(T.uyari("Bilgi", "Sorun başlığı yok", "", "Dökümde başlık sütunu yok; sorunlar issue_id ile gösterildi. README'deki sorguya başlık alanını ekleyebilirsiniz."))
    r["uyarilar"] = uyarilar
    rapor_yaz(cikti, r, uyarilar)
    (cikti.parent / "cokme_bildirimi.txt").write_text(bildirim_metni(r), encoding="utf-8")
    T.ozet_yaz(cikti.parent, PAKET, ozet(r))
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Firebase Crashlytics BigQuery dökümünden çökme ve hataları önceliklendirir.")
    p.add_argument("--girdi", nargs="+", type=Path, help="Crashlytics dökümü (.csv / .xlsx) ya da klasörü (README'deki sorguyla alınır)")
    p.add_argument("--tarih", help="Pencerenin son günü GG.AA.YYYY (varsayılan: dökümdeki son gün)")
    p.add_argument("--gun", type=int, default=7, help="Pencere uzunluğu, gün (varsayılan 7)")
    p.add_argument("--esik", type=int, default=10, help="P1 için en az etkilenen kurulum (varsayılan 10)")
    p.add_argument("--kullanicilar", type=Path, help="İsteğe bağlı: Sürüm; Aktif Kullanıcı (çökmesiz kullanıcı oranı için)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "cokme_raporu.xlsx")
    a = p.parse_args(argv)
    if not a.girdi:
        a.girdi = [ORNEK / "crashlytics_olaylari.csv"]
        a.kullanicilar = a.kullanicilar or ORNEK / "surum_kullanicilari.csv"
    bitis = T.gun(a.tarih) if a.tarih else None
    if (a.tarih and bitis is None) or a.gun < 2 or a.esik < 1:
        print("[X] --tarih GG.AA.YYYY olmalı; --gun en az 2, --esik en az 1")
        return 2
    try:
        r = calistir(a.girdi, a.cikti, bitis, a.gun, a.esik, a.kullanicilar)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    say = Counter(s["oncelik"] for s in r["sorunlar"])
    print(f"[OK] {T.tr_tarih(r['bas'])}–{T.tr_tarih(r['bitis'])} · {len(r['sorunlar'])} sorun · P1 {say['P1']} · P2 {say['P2']} · P3 {say['P3']} · en yeni sürüm {r['en_yeni']}")
    for s in r["sorunlar"]:
        if s["oncelik"] == "P1":
            print(f"[!] P1 {TUR_AD[s['tur']]}: {s['baslik']} · {s['kurulum']} kurulum · {s['neden']}")
    for u in r["uyarilar"]:
        if u["onem"] in ("Yüksek", "Orta"):
            print(f"[!] {u['tur']}: {u['kim']} · {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    print(f"[OK] Bildirim metni: {(a.cikti.parent / 'cokme_bildirimi.txt').resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
