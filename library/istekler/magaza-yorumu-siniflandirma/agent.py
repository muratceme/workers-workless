"""
Mağaza Yorumu Sınıflandırma — Gerisi Bende AI Agent (Sizden Gelenler)

1. Kod Google Play yorum dökümünü (Play Console → Yorumlar, aylık CSV; UTF-16) ve App Store yorumlarını (CSV / Excel:
   Tarih, Puan, Başlık, Yorum, Sürüm) okur. Metni olmayan yalnız puanlar modele gitmez; istatistiğe girer.
2. Model her yorumu kategori ve duyguya ayırır, varsa özellik talebini kısa bir adla yazar, yorumdan birebir kısa bir
   alıntı seçer ve cevap taslağı yazar.
3. Kod modeli denetler:
     - 1–2 yıldızlı yorum "olumlu" olamaz → "olumsuz" yapılır.
     - Alıntı yorumda birebir geçmiyorsa silinir ve işaretlenir (model uydurmasın).
     - Cevap taslağı karakter sınırını (varsayılan 350, --cevap-siniri) aşıyorsa işaretlenir.
     - Özellik talepleri aynı adla gruplanır; model önceki paketlerde kullandığı adları görür.
Kullanım:
    python agent.py                                              # örnek yorumlarla dener
    python agent.py --play reviews_reviews_com.ornek_202609.csv
    python agent.py --play play_yorumlari/ --appstore appstore_yorumlari.csv --cevap-siniri 350
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import llm
import tablo as T

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
PAKET_ADI = "magaza-yorumu-siniflandirma"
PAKET = 25
KATEGORILER = ["Abonelik / fiyat", "Hata / çökme", "Giriş / hesap", "Yapay zekâ cevap kalitesi", "Reklam", "Performans / hız",
               "Özellik isteği", "Kullanım kolaylığı / tasarım", "Bildirim", "Övgü", "Diğer"]
DUYGU = ["olumlu", "notr", "olumsuz"]
DUYGU_AD = {"olumlu": "Olumlu", "notr": "Nötr", "olumsuz": "Olumsuz"}

# Play Console yorum dökümü sütunları (support.google.com/googleplay/android-developer/answer/6135870)
PLAY = {"surum_kodu": ("App Version Code",), "surum": ("App Version Name",), "dil": ("Reviewer Language",), "cihaz": ("Device",),
        "zaman": ("Review Submit Date and Time",), "milis": ("Review Submit Millis Since Epoch",), "puan": ("Star Rating",),
        "baslik": ("Review Title",), "metin": ("Review Text",), "cevap": ("Developer Reply Text",), "baglanti": ("Review Link",)}
APPLE = {"zaman": ("Tarih", "Date", "Last Modified", "Yorum Tarihi"), "puan": ("Puan", "Rating", "Stars", "Yıldız"), "baslik": ("Başlık", "Title"),
         "metin": ("Yorum", "Review", "Body", "Content", "İçerik"), "surum": ("Sürüm", "Version", "App Version"), "dil": ("Ülke", "Territory", "Country", "Dil"),
         "cevap": ("Geliştirici Cevabı", "Developer Response", "Response")}


def _play_mi(yol: Path) -> bool:
    satirlar = T.tablo_oku(yol)
    return bool(satirlar) and "star rating" in {T.katla(c) for c in satirlar[0]}


def yorumlari_oku(play, appstore) -> list[dict]:
    yorumlar = []
    for magaza, yollar, sozluk in (("Google Play", play, PLAY), ("App Store", appstore, APPLE)):
        for y in T.dosyalar(yollar):
            if magaza == "Google Play" and not _play_mi(y):
                raise ValueError(f"{y.name} Play yorum dökümüne benzemiyor ('Star Rating' sütunu yok).")
            for r in T.kayitlar(y, sozluk, ("puan",)):
                puan = T.sayi(r["puan"])
                if puan is None or not 1 <= puan <= 5:
                    continue
                zaman = T.zaman(r.get("milis")) or T.zaman(r["zaman"])
                yorumlar.append({"id": f"{'GP' if magaza == 'Google Play' else 'AS'}{len(yorumlar) + 1}", "magaza": magaza, "zaman": zaman,
                                 "puan": int(puan), "baslik": T.metin(r["baslik"]), "metin": T.metin(r["metin"]), "surum": T.metin(r["surum"]) or "?",
                                 "dil": T.metin(r["dil"]), "cihaz": T.metin(r.get("cihaz")), "cevaplandi": bool(T.metin(r["cevap"])),
                                 "baglanti": T.metin(r.get("baglanti"))})
    return yorumlar


SEMA = {
    "type": "object",
    "properties": {"yorumlar": {"type": "array", "items": {
        "type": "object",
        "properties": {
            "id": {"type": "string"},
            "kategori": {"type": "string", "enum": KATEGORILER},
            "duygu": {"type": "string", "enum": DUYGU},
            "ozellik_talebi": {"type": "string"},
            "alinti": {"type": "string"},
            "cevap_taslagi": {"type": "string"},
        },
        "required": ["id", "kategori", "duygu", "ozellik_talebi", "alinti", "cevap_taslagi"],
        "additionalProperties": False}}},
    "required": ["yorumlar"],
    "additionalProperties": False,
}


def siniflandir(yorumlar: list[dict], urun: str, sinir: int) -> dict[str, dict]:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    sonuc, adlar = {}, []
    metinli = [y for y in yorumlar if y["metin"] or y["baslik"]]
    for i in range(0, len(metinli), PAKET):
        parca = metinli[i:i + PAKET]
        satir = [f'<yorum id="{y["id"]}" magaza="{y["magaza"]}" puan="{y["puan"]}" surum="{y["surum"]}">\n'
                 f'{llm.maskele((y["baslik"] + chr(10) if y["baslik"] else "") + y["metin"][:3000])}\n</yorum>' for y in parca]
        mesaj = "\n".join([f"<urun>{urun}</urun>", f"<cevap_siniri>{sinir}</cevap_siniri>",
                           "<bilinen_talepler>" + "; ".join(adlar) + "</bilinen_talepler>", "<yorumlar>", *satir, "</yorumlar>"])
        yanit = llm.json_iste(sistem, mesaj, SEMA)
        gecerli = {y["id"] for y in parca}
        for s in yanit.get("yorumlar", []):
            if s.get("id") in gecerli:
                sonuc[s["id"]] = s
                ad = s["ozellik_talebi"].strip()
                if ad and ad not in adlar:
                    adlar.append(ad)
    return sonuc


def _duz(s: str) -> str:
    return " ".join(s.replace("İ", "i").lower().split())


def denetle(y: dict, s: dict, sinir: int) -> list[str]:
    notlar = []
    if y["puan"] <= 2 and s["duygu"] == "olumlu":
        notlar.append(f"{y['puan']} yıldız: duygu olumlu → olumsuz (kural)")
        s["duygu"] = "olumsuz"
    if s["alinti"] and _duz(s["alinti"]) not in _duz(f"{y['baslik']} {y['metin']}"):
        notlar.append("alıntı yorumda birebir geçmiyor, silindi")
        s["alinti"] = ""
    if len(s["cevap_taslagi"]) > sinir:
        notlar.append(f"cevap {len(s['cevap_taslagi'])} karakter > {sinir}: kısaltın")
    s["ozellik_talebi"] = s["ozellik_talebi"].strip()
    return notlar


def _ort(puanlar) -> float | None:
    return round(sum(puanlar) / len(puanlar), 2) if puanlar else None


def gruplar(yorumlar, sonuc) -> list[dict]:
    g = defaultdict(list)
    ad = {}
    for y in yorumlar:
        s = sonuc.get(y["id"])
        if s and s["ozellik_talebi"]:
            k = T.katla(s["ozellik_talebi"])
            ad.setdefault(k, s["ozellik_talebi"])
            g[k].append(y)
    return sorted(({"ad": ad[k], "adet": len(v), "ortalama": _ort([y["puan"] for y in v]),
                    "magaza": ", ".join(sorted({y["magaza"] for y in v})),
                    "ornek": next((sonuc[y["id"]]["alinti"] for y in v if sonuc[y["id"]]["alinti"]), "")} for k, v in g.items()),
                  key=lambda x: (-x["adet"], x["ad"]))


def surumler(yorumlar, sonuc) -> list[dict]:
    g = defaultdict(list)
    for y in yorumlar:
        g[(y["magaza"], y["surum"])].append(y)
    satir = []
    for (m, s), v in g.items():
        olumsuz = [y for y in v if y["puan"] <= 2]
        kat = Counter(sonuc[y["id"]]["kategori"] for y in olumsuz if y["id"] in sonuc)
        satir.append({"magaza": m, "surum": s, "adet": len(v), "ortalama": _ort([y["puan"] for y in v]), "olumsuz": len(olumsuz),
                      "oran": len(olumsuz) / len(v), "konu": kat.most_common(1)[0][0] if kat else ""})
    return sorted(satir, key=lambda x: (x["magaza"], x["surum"]))


AI = T.MOR


def rapor_yaz(cikti: Path, yorumlar, sonuc, notlar, sinir: int) -> None:
    wb = T.yeni_kitap()
    bas = ["No", "Mağaza", "Zaman", "Puan", "Sürüm", "Dil / Ülke", "Başlık", "Yorum", "Kategori", "Duygu", "Özellik Talebi", "Alıntı",
           "Cevap Taslağı", "Karakter", "Zaten Cevaplandı", "Kod Notu", "Bağlantı", "Onay"]
    gen = (8, 12, 16, 6, 10, 9, 22, 60, 22, 9, 24, 34, 60, 9, 10, 34, 20, 8)
    sirali = sorted(yorumlar, key=lambda y: (y["zaman"] is None, -(y["zaman"].timestamp() if y["zaman"] else 0)))

    def yaz(ws, y):
        s = sonuc.get(y["id"])
        if not s:
            T.satir(ws, [y["id"], y["magaza"], y["zaman"], y["puan"], y["surum"], y["dil"], y["baslik"], y["metin"] or "(yalnız puan)", "", "", "", "", "", None,
                         "evet" if y["cevaplandi"] else "", "", y["baglanti"]], bicim={3: T.ZF}, sar=(8,))
            return
        T.satir(ws, [y["id"], y["magaza"], y["zaman"], y["puan"], y["surum"], y["dil"], y["baslik"], y["metin"], s["kategori"], DUYGU_AD[s["duygu"]],
                     s["ozellik_talebi"], s["alinti"], s["cevap_taslagi"], len(s["cevap_taslagi"]), "evet" if y["cevaplandi"] else "",
                     "; ".join(notlar.get(y["id"], [])), y["baglanti"], ""],
                bicim={3: T.ZF}, sar=(8, 12, 13, 16),
                dolgu={4: T.KIRMIZI if y["puan"] <= 2 else None, 9: AI, 10: AI, 11: AI, 12: AI, 13: AI,
                       14: T.KIRMIZI if len(s["cevap_taslagi"]) > sinir else None, 18: T.SARI})

    ws = T.sayfa(wb, "Olumsuz", bas, gen, "B2")
    for y in sirali:
        if y["puan"] <= 2 or sonuc.get(y["id"], {}).get("duygu") == "olumsuz":
            yaz(ws, y)
    ws = T.sayfa(wb, "Özellik Talepleri", ["Talep", "Yorum Sayısı", "Ortalama Puan", "Mağaza", "Örnek Alıntı"], (34, 12, 13, 22, 60), "A2")
    for g in gruplar(yorumlar, sonuc):
        T.satir(ws, [g["ad"], g["adet"], g["ortalama"], g["magaza"], g["ornek"]], bicim={3: "0.00"}, sar=(5,))
    ws = T.sayfa(wb, "Yorumlar", bas, gen, "B2")
    for y in sirali:
        yaz(ws, y)
    ws = T.sayfa(wb, "Sürümler", ["Mağaza", "Sürüm", "Yorum", "Ortalama Puan", "1–2 Yıldız", "1–2 Yıldız Oranı", "En Sık Olumsuz Konu"], (12, 12, 8, 13, 10, 14, 28), "A2")
    for s in surumler(yorumlar, sonuc):
        T.satir(ws, [s["magaza"], s["surum"], s["adet"], s["ortalama"], s["olumsuz"], s["oran"], s["konu"]], bicim={4: "0.00", 6: "0.0%"})
    o = T.sayfa(wb, "Özet", ["Gösterge", "Değer"], (36, 40))
    for m in ("Google Play", "App Store"):
        v = [y for y in yorumlar if y["magaza"] == m]
        if v:
            T.satir(o, [f"{m}: yorum / ortalama puan", f"{len(v)} / {_ort([y['puan'] for y in v]):.2f}"])
    T.satir(o, ["Metinli yorum (modele giden)", sum(1 for y in yorumlar if y["metin"] or y["baslik"])])
    T.satir(o, ["1–2 yıldız", sum(y["puan"] <= 2 for y in yorumlar)])
    T.satir(o, ["Cevaplanmamış 1–2 yıldız", sum(y["puan"] <= 2 and not y["cevaplandi"] for y in yorumlar)])
    T.satir(o, ["Kod notu olan yorum", len([n for n in notlar.values() if n])])
    T.satir(o, [])
    for k, n in Counter(s["kategori"] for s in sonuc.values()).most_common():
        T.satir(o, [k, n])
    T.satir(o, [])
    T.satir(o, ["Model", llm.kullanim_ozeti()])
    T.satir(o, ["Renkler", "Mor: yapay zekâ; kırmızı puan: 1–2 yıldız; kırmızı karakter: cevap sınırı aşıldı"])
    T.kaydet(wb, cikti)


def ozet(yorumlar, sonuc) -> dict:
    gunluk = defaultdict(list)
    for y in yorumlar:
        if y["zaman"]:
            gunluk[y["zaman"].date().isoformat()].append(y)
    return {"gunluk": {g: {"yorum": len(v), "ortalama_puan": _ort([y["puan"] for y in v]), "olumsuz": sum(y["puan"] <= 2 for y in v)}
                       for g, v in sorted(gunluk.items())},
            "kategoriler": dict(Counter(s["kategori"] for s in sonuc.values()).most_common()),
            "ozellik_talepleri": [{"ad": g["ad"], "adet": g["adet"]} for g in gruplar(yorumlar, sonuc)[:10]]}


def calistir(play, appstore, cikti: Path, urun: str = "abonelikli mobil uygulama", sinir: int = 350, evet: bool = False) -> dict:
    yorumlar = yorumlari_oku(play or [], appstore or [])
    if not yorumlar:
        raise llm.LLMHatasi("Okunacak yorum yok.")
    metinli = sum(1 for y in yorumlar if y["metin"] or y["baslik"])
    print(f"[OK] {len(yorumlar)} yorum · metinli {metinli} · 1–2 yıldız {sum(y['puan'] <= 2 for y in yorumlar)}")
    llm.onay_al(f"{metinli} yorumun başlık ve metni (puan ve sürümle; e-posta / telefon maskeli) gönderilecek.", evet)
    sonuc = siniflandir(yorumlar, urun, sinir)
    notlar = {y["id"]: denetle(y, sonuc[y["id"]], sinir) for y in yorumlar if y["id"] in sonuc}
    rapor_yaz(cikti, yorumlar, sonuc, notlar, sinir)
    T.ozet_yaz(cikti.parent, PAKET_ADI, ozet(yorumlar, sonuc))
    return {"yorumlar": yorumlar, "sonuc": sonuc, "notlar": notlar, "gruplar": gruplar(yorumlar, sonuc)}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    p = argparse.ArgumentParser(description="Google Play ve App Store yorumlarını sınıflandırır, özellik taleplerini gruplar ve cevap taslağı yazar.")
    p.add_argument("--play", nargs="+", type=Path, help="Play Console yorum dökümü (.csv, UTF-16) ya da klasörü")
    p.add_argument("--appstore", nargs="+", type=Path, help="App Store yorumları (.csv / .xlsx: Tarih, Puan, Başlık, Yorum, Sürüm)")
    p.add_argument("--urun", default="abonelikli mobil uygulama", help="Modele verilecek kısa ürün tanımı")
    p.add_argument("--cevap-siniri", type=int, default=350, help="Cevap taslağı için karakter sınırı (varsayılan 350)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "magaza_yorumlari.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    if not a.play and not a.appstore:
        a.play, a.appstore = [ORNEK / "reviews_reviews_com.ornek.uygulama_202609.csv"], [ORNEK / "appstore_yorumlari.csv"]
    if a.cevap_siniri < 50:
        print("[X] --cevap-siniri en az 50 olmalı")
        return 2
    try:
        s = calistir(a.play, a.appstore, a.cikti, a.urun, a.cevap_siniri, a.evet)
    except (llm.LLMHatasi, ValueError) as h:
        print(f"[X] {h}")
        return 1
    say = Counter(x["duygu"] for x in s["sonuc"].values())
    print(f"[OK] {len(s['sonuc'])} yorum sınıflandı · olumlu {say['olumlu']} · nötr {say['notr']} · olumsuz {say['olumsuz']}")
    for g in s["gruplar"][:5]:
        print(f"     özellik talebi: {g['ad']} ({g['adet']})")
    n = sum(1 for x in s["notlar"].values() if x)
    if n:
        print(f"[!] {n} yorumda kod notu var (Kod Notu sütunu)")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
