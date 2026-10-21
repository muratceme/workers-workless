"""
Uygulama Abonelik Gelir Raporu — Gerisi Bende kod bloğu (Sizden Gelenler)

Google Play kazanç raporu ile App Store Connect özet satış raporunu tek tabloda birleştirir:
  Google Play  "Transaction Type" sütununa göre:
                 Charge            → satış (brüt)         Charge refund     → iade
                 Google fee        → mağaza komisyonu     Google fee refund → komisyon iadesi
                 diğer türler      → "Diğer" (adı korunur, uyarı verilir)
               Tutarlar "Amount (Merchant Currency)" sütunundan alınır (satıcı para birimi).
               Net = o günün bütün satırlarının toplamı (satış − iade − komisyon + komisyon iadesi + diğer).
  App Store    Satır tutarı = Units × Developer Proceeds (birim başına geliştirici payı; vergi ve Apple komisyonu düşülmüş).
               İadede Units negatiftir, böylece tutar da negatif olur. Apple raporu komisyonu ve vergiyi ayrı vermediği
               için App Store'da brüt ve komisyon hesaplanmaz; net gelir gösterilir.
               "Subscription" sütunu: New → yeni abonelik, Renewal → yenileme.
  Para birimi  Tüm net tutarlar --para-birimi'ne (varsayılan: Play satıcı para birimi) --kur dosyasıyla çevrilir.
               Kuru verilmeyen para birimi toplama katılmaz ve uyarı verilir.
  İptaller     İsteğe bağlı abonelik olayları dosyasından (Tarih, Olay, Adet) sayılır.
İnternete bağlanmaz.
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path

import tablo as T

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
PAKET = "uygulama-abonelik-gelir-raporu"
SIFIR = Decimal(0)

PLAY = {"siparis": ("Description",), "tarih": ("Transaction Date",), "tur": ("Transaction Type",), "urun": ("Product Title",),
        "sku": ("SKU ID",), "plan": ("Base Plan or Purchase Option ID",), "ulke": ("Buyer Country",),
        "para": ("Merchant Currency",), "tutar": ("Amount (Merchant Currency)",)}
APPLE = {"sku": ("SKU",), "baslik": ("Title",), "tur_kodu": ("Product Type Identifier",), "adet": ("Units",),
         "pay": ("Developer Proceeds", "Developer Proceeds (per unit)"), "tarih": ("Begin Date",), "ulke": ("Country Code",),
         "para": ("Currency of Proceeds",), "abonelik": ("Subscription",), "ust": ("Parent Identifier",), "donem": ("Period",)}
OLAY = {"tarih": ("Tarih", "Event Date", "Date"), "olay": ("Olay", "Event", "Olay Türü"), "adet": ("Adet", "Quantity", "Sayı"),
        "platform": ("Platform", "Mağaza"), "urun": ("Ürün", "Abonelik", "Subscription Name", "Product")}
KUR = {"para": ("Para Birimi", "Currency", "Döviz"), "kur": ("Kur", "Rate", "Oran")}
PLAY_TUR = {"charge": "Satış", "charge refund": "İade", "google fee": "Komisyon", "google fee refund": "Komisyon iadesi"}


def _olay_turu(s: str) -> str:
    k = T.katla(s)
    if any(x in k for x in ("iptal", "cancel")):
        return "İptal"
    if any(x in k for x in ("iade", "refund")):
        return "İade"
    if any(x in k for x in ("yenile", "renew")):
        return "Yenileme"
    if any(x in k for x in ("yeni", "start", "subscribe", "new")):
        return "Yeni"
    return "Diğer"


def oku_play(yollar, uyarilar) -> list[dict]:
    islemler = []
    for r in T.hepsini_oku(yollar, PLAY, ("tarih", "tur", "tutar")):
        g, tutar = T.gun(r["tarih"]), T.sayi(r["tutar"], ".")
        if g is None or tutar is None:
            uyarilar.append(T.uyari("Orta", "Okunamayan Play satırı", f"{r['_dosya']} satır {r['_satir']}", "Tarih ya da tutar okunamadı; satır atlandı."))
            continue
        tur = PLAY_TUR.get(T.katla(r["tur"]))
        if tur is None:
            tur = "Diğer"
            uyarilar.append(T.uyari("Bilgi", "Tanınmayan Play işlem türü", T.metin(r["tur"]), "Net toplama katıldı, satış / iade / komisyon sayılmadı."))
        islemler.append({"platform": "Google Play", "tarih": g, "tur": tur, "ham_tur": T.metin(r["tur"]), "siparis": T.metin(r["siparis"]),
                         "urun": T.metin(r["sku"]) or T.metin(r["urun"]), "urun_adi": T.metin(r["urun"]), "ulke": T.metin(r["ulke"]),
                         "para": T.metin(r["para"]).upper(), "tutar": tutar, "adet": 1 if tur in ("Satış", "İade") else 0, "abonelik": ""})
    return islemler


def oku_apple(yollar, uyarilar) -> list[dict]:
    islemler = []
    for r in T.hepsini_oku(yollar, APPLE, ("adet", "pay", "tarih", "para")):
        g, adet, pay = T.gun(r["tarih"], abd=True), T.sayi(r["adet"], "."), T.sayi(r["pay"], ".")
        if g is None or adet is None or pay is None:
            uyarilar.append(T.uyari("Orta", "Okunamayan App Store satırı", f"{r['_dosya']} satır {r['_satir']}", "Tarih, adet ya da geliştirici payı okunamadı; satır atlandı."))
            continue
        if adet == 0 or pay == 0:
            continue  # ücretsiz indirme / güncelleme: gelir yok
        if adet < 0 and pay < 0:
            uyarilar.append(T.uyari("Orta", "Beklenmeyen iade işareti", f"{r['_dosya']} satır {r['_satir']}",
                                    "Units ve Developer Proceeds ikisi de negatif; Apple'ın belgelediği biçimde iadede yalnız Units negatiftir. Raporu kontrol edin."))
        abonelik = T.metin(r["abonelik"])
        islemler.append({"platform": "App Store", "tarih": g, "tur": "İade" if adet < 0 else "Satış", "ham_tur": T.metin(r["tur_kodu"]), "siparis": "",
                         "urun": T.metin(r["sku"]) or T.metin(r["baslik"]), "urun_adi": T.metin(r["baslik"]), "ulke": T.metin(r["ulke"]),
                         "para": T.metin(r["para"]).upper(), "tutar": adet * pay, "adet": int(abs(adet)), "abonelik": abonelik})
    return islemler


def oku_kur(yol: Path | None) -> dict:
    if not yol:
        return {}
    kur = {}
    for r in T.kayitlar(yol, KUR, ("para", "kur")):
        k = T.sayi(r["kur"])
        if T.metin(r["para"]) and k:
            kur[T.metin(r["para"]).upper()] = k
    return kur


def oku_olaylar(yol: Path | None, uyarilar) -> list[dict]:
    if not yol:
        return []
    olaylar = []
    for r in T.kayitlar(yol, OLAY, ("tarih", "olay")):
        g = T.gun(r["tarih"])
        if g is None:
            uyarilar.append(T.uyari("Orta", "Okunamayan olay satırı", f"satır {r['_satir']}", "Tarih okunamadı; atlandı."))
            continue
        olaylar.append({"tarih": g, "olay": _olay_turu(r["olay"]), "ham": T.metin(r["olay"]), "adet": int(T.sayi(r["adet"]) or 1),
                        "platform": T.metin(r["platform"]) or "—", "urun": T.metin(r["urun"])})
    return olaylar


def _bos():
    return {"satis_adet": 0, "iade_adet": 0, "brut": SIFIR, "iade": SIFIR, "komisyon": SIFIR, "diger": SIFIR, "net": SIFIR, "yeni": 0, "yenileme": 0}


def analiz(islemler: list[dict], hedef: str, kur: dict, olaylar: list[dict] | None = None) -> dict:
    uyarilar, eksik_kur = [], defaultdict(Decimal)
    gunluk = defaultdict(_bos)          # (tarih, platform)
    urunler = defaultdict(_bos)         # (platform, ürün)
    ulkeler = defaultdict(_bos)         # (platform, ülke)
    for i in islemler:
        oran = Decimal(1) if i["para"] in ("", hedef) else kur.get(i["para"])
        if oran is None:
            eksik_kur[i["para"]] += i["tutar"]
            i["cevrilmis"] = None
            continue
        t = T.yuvarla(i["tutar"] * oran)
        i["cevrilmis"] = t
        for anahtar, kova in (((i["tarih"], i["platform"]), gunluk), ((i["platform"], i["urun"]), urunler), ((i["platform"], i["ulke"] or "—"), ulkeler)):
            k = kova[anahtar]
            k["net"] += t
            if i["tur"] == "Satış":
                k["satis_adet"] += i["adet"]
                k["brut"] += t if i["platform"] == "Google Play" else SIFIR
                k["yeni"] += i["adet"] if T.katla(i["abonelik"]) == "new" else 0
                k["yenileme"] += i["adet"] if T.katla(i["abonelik"]) == "renewal" else 0
            elif i["tur"] == "İade":
                k["iade_adet"] += i["adet"]
                k["iade"] += t
            elif i["tur"] in ("Komisyon", "Komisyon iadesi"):
                k["komisyon"] += t
            else:
                k["diger"] += t
    for para, tutar in sorted(eksik_kur.items()):
        uyarilar.append(T.uyari("Yüksek", "Kuru verilmeyen para birimi", para or "(boş)",
                                f"Toplam {T.yuvarla(tutar)} {para} toplamlara katılmadı. --kur dosyasına {para} için {hedef} karşılığını ekleyin."))
    iptal = defaultdict(int)
    for o in olaylar or []:
        if o["olay"] == "İptal":
            iptal[(o["tarih"], o["platform"])] += o["adet"]
    toplam = _bos()
    platform_net = defaultdict(lambda: SIFIR)
    for (g, p), k in gunluk.items():
        platform_net[p] += k["net"]
        for a in toplam:
            toplam[a] += k[a]
    toplam["iptal"] = sum(iptal.values())
    play_brut = sum((k["brut"] for (g, p), k in gunluk.items() if p == "Google Play"), SIFIR)
    play_kom = sum((k["komisyon"] for (g, p), k in gunluk.items() if p == "Google Play"), SIFIR)
    toplam["komisyon_orani"] = (-play_kom / play_brut) if play_brut else None
    toplam["iade_orani"] = (Decimal(toplam["iade_adet"]) / toplam["satis_adet"]) if toplam["satis_adet"] else None
    return {"gunluk": gunluk, "urunler": urunler, "ulkeler": ulkeler, "iptal": iptal, "toplam": toplam, "platform_net": dict(platform_net),
            "uyarilar": uyarilar, "hedef": hedef}


def _ay(d: date) -> str:
    return d.strftime("%Y-%m")


def rapor_yaz(cikti: Path, r: dict, islemler: list[dict], olaylar: list[dict]) -> None:
    h, para = r["hedef"], f"Tutar ({r['hedef']})"
    wb = T.yeni_kitap()
    t = r["toplam"]
    ws = T.sayfa(wb, "Özet", ["Gösterge", "Değer", "Açıklama"], (34, 18, 90))
    satirlar = [("Net gelir (tüm platformlar)", t["net"], T.PF, f"{h}; mağaza payı ve iadeler düşülmüş"),
                *[(f"Net gelir · {p}", v, T.PF, "") for p, v in sorted(r["platform_net"].items())],
                ("Brüt satış · Google Play", sum((k["brut"] for (g, p), k in r["gunluk"].items() if p == "Google Play"), SIFIR), T.PF, "Charge satırları"),
                ("İade tutarı", t["iade"], T.PF, "Play: Charge refund; App Store: negatif Units satırları (geliştirici payı)"),
                ("Mağaza komisyonu · Google Play", t["komisyon"], T.PF, "Google fee + Google fee refund"),
                ("Efektif komisyon oranı · Google Play", t["komisyon_orani"], T.YF, "komisyon ÷ brüt satış"),
                ("Satış adedi", t["satis_adet"], T.TAMF, ""), ("İade adedi", t["iade_adet"], T.TAMF, ""),
                ("İade oranı (adet)", t["iade_orani"], T.YF, "iade adedi ÷ satış adedi"),
                ("Yeni abonelik · App Store", t["yeni"], T.TAMF, "Subscription = New"), ("Yenileme · App Store", t["yenileme"], T.TAMF, "Subscription = Renewal"),
                ("İptal (abonelik olayları)", t["iptal"] if olaylar else None, T.TAMF, "olay dosyası verilmediyse boş")]
    for ad, deger, bicim, aciklama in satirlar:
        T.satir(ws, [ad, deger, aciklama], bicim={2: bicim})
    T.satir(ws, [])
    T.satir(ws, ["Not", "", "App Store raporu komisyon ve vergiyi ayrı vermez; App Store için yalnız net (geliştirici payı) gösterilir."], sar=(3,))

    g = T.sayfa(wb, "Günlük", ["Tarih", "Platform", "Satış Adedi", "İade Adedi", "Brüt (Play)", "İade", "Komisyon (Play)", "Diğer", "Net", "İptal"],
                (12, 14, 11, 11, 15, 15, 15, 12, 15, 9))
    for (d, p), k in sorted(r["gunluk"].items()):
        T.satir(g, [d, p, k["satis_adet"], k["iade_adet"], k["brut"] if p == "Google Play" else None, k["iade"], k["komisyon"] if p == "Google Play" else None,
                    k["diger"], k["net"], r["iptal"].get((d, p), 0)], bicim={1: T.TF, 5: T.PF, 6: T.PF, 7: T.PF, 8: T.PF, 9: T.PF})

    aylik = defaultdict(_bos)
    for (d, p), k in r["gunluk"].items():
        for a in k:
            aylik[(_ay(d), p)][a] += k[a]
    a = T.sayfa(wb, "Aylık", ["Ay", "Platform", "Satış Adedi", "İade Adedi", "Brüt (Play)", "İade", "Komisyon (Play)", "Net", "İptal"], (10, 14, 11, 11, 15, 15, 15, 15, 9))
    for (ay, p), k in sorted(aylik.items()):
        ipt = sum(n for (d, pp), n in r["iptal"].items() if _ay(d) == ay and pp == p)
        T.satir(a, [ay, p, k["satis_adet"], k["iade_adet"], k["brut"] if p == "Google Play" else None, k["iade"], k["komisyon"] if p == "Google Play" else None,
                    k["net"], ipt], bicim={5: T.PF, 6: T.PF, 7: T.PF, 8: T.PF})

    u = T.sayfa(wb, "Ürünler", ["Platform", "Ürün / SKU", "Satış Adedi", "İade Adedi", "İade", "Net"], (14, 34, 11, 11, 15, 15))
    for (p, urun), k in sorted(r["urunler"].items(), key=lambda x: -x[1]["net"]):
        T.satir(u, [p, urun, k["satis_adet"], k["iade_adet"], k["iade"], k["net"]], bicim={5: T.PF, 6: T.PF})
    ul = T.sayfa(wb, "Ülkeler", ["Platform", "Ülke", "Satış Adedi", "İade Adedi", "Net"], (14, 10, 11, 11, 15))
    for (p, ulke), k in sorted(r["ulkeler"].items(), key=lambda x: -x[1]["net"]):
        T.satir(ul, [p, ulke, k["satis_adet"], k["iade_adet"], k["net"]], bicim={5: T.PF})
    if olaylar:
        o = T.sayfa(wb, "Abonelik Olayları", ["Tarih", "Platform", "Ürün", "Olay", "Kaynaktaki Ad", "Adet"], (12, 14, 30, 12, 24, 8))
        for x in sorted(olaylar, key=lambda x: (x["tarih"], x["platform"])):
            T.satir(o, [x["tarih"], x["platform"], x["urun"], x["olay"], x["ham"], x["adet"]], bicim={1: T.TF})
    i = T.sayfa(wb, "İşlemler", ["Tarih", "Platform", "Tür", "Kaynaktaki Tür", "Sipariş", "Ürün / SKU", "Ülke", "Para Birimi", "Tutar", para, "Abonelik"],
                (12, 14, 16, 22, 26, 30, 7, 8, 13, 15, 10))
    for x in sorted(islemler, key=lambda x: (x["tarih"], x["platform"], x["siparis"])):
        T.satir(i, [x["tarih"], x["platform"], x["tur"], x["ham_tur"], x["siparis"], x["urun"], x["ulke"], x["para"], x["tutar"], x.get("cevrilmis"), x["abonelik"]],
                bicim={1: T.TF, 9: T.PF, 10: T.PF}, dolgu={10: T.KIRMIZI if x.get("cevrilmis") is None else None})
    T.uyari_sayfasi(wb, r["uyarilar"])
    T.kaydet(wb, cikti)


def ozet(r: dict) -> dict:
    gunluk = defaultdict(lambda: {"net": SIFIR, "brut_play": SIFIR, "iade": SIFIR, "komisyon_play": SIFIR, "satis_adet": 0, "iade_adet": 0, "iptal": 0})
    for (d, p), k in r["gunluk"].items():
        x = gunluk[d.isoformat()]
        x["net"] += k["net"]
        x["iade"] += k["iade"]
        x["satis_adet"] += k["satis_adet"]
        x["iade_adet"] += k["iade_adet"]
        if p == "Google Play":
            x["brut_play"] += k["brut"]
            x["komisyon_play"] += k["komisyon"]
    for (d, p), n in r["iptal"].items():
        gunluk[d.isoformat()]["iptal"] += n
    return {"para_birimi": r["hedef"], "gunluk": dict(sorted(gunluk.items()))}


def calistir(play, apple, cikti: Path, hedef: str | None = None, kur_yol: Path | None = None, olay_yol: Path | None = None) -> dict:
    uyarilar = []
    islemler = oku_play(play, uyarilar) + oku_apple(apple, uyarilar)
    if not islemler:
        raise ValueError("Okunacak gelir satırı yok: --play ve / veya --apple ile rapor dosyası verin.")
    play_para = sorted({i["para"] for i in islemler if i["platform"] == "Google Play" and i["para"]})
    hedef = (hedef or (play_para[0] if play_para else "TRY")).upper()
    olaylar = oku_olaylar(olay_yol, uyarilar)
    r = analiz(islemler, hedef, oku_kur(kur_yol), olaylar)
    r["uyarilar"] = uyarilar + r["uyarilar"]
    if r["toplam"]["iade_orani"] is not None and r["toplam"]["iade_orani"] > Decimal("0.1"):
        r["uyarilar"].append(T.uyari("Orta", "Yüksek iade oranı", "", f"İade adedi satışların %{T.yuvarla(r['toplam']['iade_orani'] * 100, '0.1')}'i."))
    rapor_yaz(cikti, r, islemler, olaylar)
    T.ozet_yaz(cikti.parent, PAKET, ozet(r))
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Google Play kazanç ve App Store Connect satış raporlarından abonelik gelir raporu çıkarır.")
    p.add_argument("--play", nargs="+", type=Path, help="Google Play kazanç raporu (earnings_YYYYMM.zip / .csv) ya da klasörü")
    p.add_argument("--apple", nargs="+", type=Path, help="App Store Connect özet satış raporu (.txt / .gz / .csv) ya da klasörü")
    p.add_argument("--para-birimi", help="Raporlama para birimi (varsayılan: Play satıcı para birimi, o da yoksa TRY)")
    p.add_argument("--kur", type=Path, help="Kur dosyası: Para Birimi; Kur (1 birimin raporlama para birimindeki karşılığı)")
    p.add_argument("--olaylar", type=Path, help="İsteğe bağlı abonelik olayları: Tarih, Olay (Yeni / Yenileme / İptal / İade), Adet, Platform, Ürün")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "abonelik_gelir_raporu.xlsx")
    a = p.parse_args(argv)
    if not a.play and not a.apple:
        a.play, a.apple = [ORNEK / "play_earnings_202609.csv"], [ORNEK / "appstore_satis_202609.txt"]
        a.kur = a.kur or ORNEK / "kur.csv"
        a.olaylar = a.olaylar or ORNEK / "abonelik_olaylari.csv"
    try:
        r = calistir(a.play, a.apple, a.cikti, a.para_birimi, a.kur, a.olaylar)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    t, h = r["toplam"], r["hedef"]
    print(f"[OK] Net gelir {T.yuvarla(t['net'])} {h} · " + " · ".join(f"{p} {T.yuvarla(v)}" for p, v in sorted(r["platform_net"].items())))
    print(f"[OK] Satış {t['satis_adet']} · iade {t['iade_adet']} · Play komisyonu {T.yuvarla(t['komisyon'])} {h}" +
          (f" (brütün %{T.yuvarla(t['komisyon_orani'] * 100, '0.1')})" if t["komisyon_orani"] is not None else "") + f" · iptal {t['iptal']}")
    for u in r["uyarilar"]:
        if u["onem"] in ("Yüksek", "Orta"):
            print(f"[!] {u['tur']}: {u['kim']} · {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
