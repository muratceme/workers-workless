"""
Ürün Başlığı SEO Kontrolü — Workers / Workless kod bloğu
Kategori ve İçerik Yönetimi › E-ticaret İçerik Uzmanı
Ürün başlıklarını uzunluk, marka - ürün tipi - özellik sırası ve yasaklı ifadeler açısından kontrol edip öneri sunar.
  Uzunluk           = --min (30) ile --max (100) karakter arası
  Sıra              = başlık markayla başlar, ardından ürün tipi gelir
  Özellik           = Renk, Materyal, Beden / Ebat, Model sütunlarına yazılan değer başlıkta geçer
  Yasaklı ifade     = varsayılan liste ya da --yasakli dosyası (+ --ek-yasakli)
  Biçim             = harflerin yarısından fazlası büyük harf · ! ? * # gibi özel karakter ya da emoji · aynı kelime iki kez ·
                      fazla boşluk · aynı başlık başka üründe
  Önerilen başlık   = sıra doğruysa başlığın kendisi, değilse Marka + Ürün tipi + diğer kelimeler; yasaklı ifade, özel karakter
                      ve tekrar ayıklanır, başlıkta olmayan özellikler sona eklenir
"""
from __future__ import annotations

from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
BASLIK_DOLGU, BASLIK_YAZI = PatternFill("solid", fgColor="1D1D1F"), Font(bold=True, color="FFFFFF")
RENK = {"Yüksek": "FDE2E1", "Orta": "FFF4CE", "Bilgi": "E8F0FE"}
KIRMIZI, SARI, MAVI, YESIL = "FDE2E1", "FFF4CE", "E8F0FE", "E3F4E1"
UST = Alignment(vertical="top", wrap_text=True)
PF, TF = "#,##0.00", "DD.MM.YYYY"
KURUS = Decimal("0.01")


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(s or "").translate(_TR).lower()).split())


def tarih(x) -> date | None:
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    s = str(x or "").strip()[:10]
    for f in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def para(x) -> Decimal | None:
    """1.250,50 · 1250.50 · 750.000 · %12,5 yazımlarını Decimal'e çevirir."""
    if x in (None, ""):
        return None
    if isinstance(x, (int, float, Decimal)):
        return Decimal(str(x))
    s = re.sub(r"[^\d,.\-]", "", str(x))
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def yuvarla(x: Decimal, hane: str = "0.01") -> Decimal:
    return x.quantize(Decimal(hane), rounding=ROUND_HALF_UP)


def metin(x) -> str:
    return str(x if x is not None else "").strip()


def tablo_oku(yol: Path) -> list[list]:
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, data_only=True, read_only=True)
        try:
            satirlar = [list(r) for r in wb.active.iter_rows(values_only=True)]
        finally:
            wb.close()
    else:
        icerik = None
        for kod in ("utf-8-sig", "cp1254"):
            try:
                icerik = yol.read_text(encoding=kod)
                break
            except UnicodeDecodeError:
                continue
        ilk = "\n".join(icerik.splitlines()[:10])
        satirlar = list(csv.reader(icerik.splitlines(), delimiter=max(";\t", key=ilk.count)))
    return [r for r in satirlar if any(c not in (None, "") for c in r)]


def kayitlar(yol: Path, sozluk: dict, zorunlu: tuple[str, ...]) -> list[dict]:
    """Başlık satırını bulur (ilk 15 satırda); sütun adlarını Türkçe karakter ve büyük-küçük harf ayırmadan eşleştirir."""
    s = tablo_oku(yol)
    for bi, r in enumerate(s[:15]):
        b = [katla(c) for c in r]
        es = {}
        for alan, adlar in sozluk.items():
            j = next((b.index(a) for a in adlar if a in b and b.index(a) not in es.values()), None)
            if j is not None:
                es[alan] = j
        if all(z in es for z in zorunlu):
            return [{a: None for a in sozluk} | {a: (r[j] if j < len(r) else None) for a, j in es.items()} | {"_satir": n}
                    for n, r in enumerate(s[bi + 1:], bi + 2)]
    raise ValueError(f"{yol.name}: başlık satırı bulunamadı; gerekli sütunlar: {', '.join(sozluk[z][0] for z in zorunlu)}")


def uyari(onem: str, tur: str, kim, aciklama: str) -> dict:
    return {"onem": onem, "tur": tur, "kim": metin(kim), "aciklama": aciklama}


def sirala(uyarilar: list[dict]) -> list[dict]:
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    return sorted(uyarilar, key=lambda u: (sira[u["onem"]], u["tur"], u["kim"]))


def _baslik(ws, basliklar, genislik=(), dondur="A2"):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = dondur


def _satir(ws, degerler, para_s=(), tarih_s=(), dolgu=None, sar=()):
    """Satır ekler; para_s / tarih_s sütunlarına biçim, dolgu = {sütun: renk}, sar = metni kaydırılacak sütunlar."""
    ws.append([float(d) if isinstance(d, Decimal) else d for d in degerler])
    n = ws.max_row
    for j in para_s:
        ws.cell(n, j).number_format = PF
    for j in tarih_s:
        ws.cell(n, j).number_format = TF
    for j, renk in (dolgu or {}).items():
        if renk:
            ws.cell(n, j).fill = PatternFill("solid", fgColor=renk)
    for j in sar:
        ws.cell(n, j).alignment = UST
    return n


def uyari_sayfasi(wb, uyarilar: list[dict]) -> None:
    ws = wb.create_sheet("Uyarılar")
    _baslik(ws, ["Önem", "Tür", "Kim / Ne", "Açıklama"], (9, 32, 28, 100))
    for u in uyarilar:
        _satir(ws, [u["onem"], u["tur"], u["kim"], u["aciklama"]], dolgu={1: RENK[u["onem"]]}, sar=(4,))


def kaydet(wb, cikti: Path) -> None:
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)

SUTUN = {"kod": ("urun kodu", "stok kodu", "barkod", "kod", "sku", "model kodu"), "baslik": ("baslik", "urun basligi", "urun adi", "urun ismi"), "marka": ("marka", "marka adi"),
         "tip": ("urun tipi", "urun turu", "tip", "kategori"), "renk": ("renk",), "materyal": ("materyal", "malzeme", "kumas"), "ebat": ("beden", "ebat", "boyut", "olcu", "hacim"),
         "model": ("model", "seri")}
OZELLIK = (("renk", "Renk"), ("materyal", "Materyal"), ("ebat", "Beden / ebat"), ("model", "Model"))
YASAKLI = ("en ucuz", "en iyi", "en kaliteli", "1 numara", "bir numara", "indirim", "indirimli", "kampanya", "kampanyali", "ucretsiz kargo", "bedava", "firsat", "son fiyat", "stokta",
           "hemen al", "acele", "sinirli stok", "sok fiyat")
OZEL = re.compile(r"[!?*#$@~^|<>{}\[\]_=+]|[☀-➿\U0001F300-\U0001FAFF]")
_KUCUK = str.maketrans("Iİ", "ıi")
_BUYUK = str.maketrans("ıi", "Iİ")


def tr_baslik(kelime: str) -> str:
    k = kelime.translate(_KUCUK).lower()
    return k[:1].translate(_BUYUK).upper() + k[1:] if k else k


def oku(yol: Path) -> tuple[list[dict], list[dict]]:
    uy, l = [], []
    for r in kayitlar(yol, SUTUN, ("kod", "baslik")):
        kod = metin(r["kod"])
        if not kod:
            continue
        ham = "" if r["baslik"] is None else str(r["baslik"])
        if not ham.strip():
            uy.append(uyari("Yüksek", "Başlık boş", kod, f"Satır {r['_satir']}: başlık yazılmamış. Ürün alınmadı."))
            continue
        l.append({"kod": kod, "ham": ham, "baslik": " ".join(ham.split()), **{a: metin(r[a]) for a in ("marka", "tip", "renk", "materyal", "ebat", "model")}})
    return l, uy


def yasakli_oku(yol: Path) -> list[str]:
    return [katla(r[0]) for r in tablo_oku(yol) if r and katla(r[0]) and katla(r[0]) not in ("yasakli ifade", "ifade", "kelime")]


def _bul(kelimeler: list[str], ifade: str, ek: bool = False) -> int | None:
    """Katlanmış kelime listesinde ifadenin başladığı sıra; ek=True ise kelimenin ek almış hâli de sayılır (pamuk → pamuklu)."""
    p = katla(ifade).split()
    if not p:
        return None
    esit = (lambda a, b: a.startswith(b)) if ek else (lambda a, b: a == b)
    return next((i for i in range(len(kelimeler) - len(p) + 1) if all(esit(a, b) for a, b in zip(kelimeler[i:i + len(p)], p))), None)


def kontrol(urunler: list[dict], en_az: int = 30, en_cok: int = 100, yasakli: tuple[str, ...] | list[str] = YASAKLI) -> dict:
    if not urunler:
        raise ValueError("Kontrol edilecek başlık yok.")
    yasakli = sorted({katla(y) for y in yasakli if katla(y)}, key=lambda y: (-len(y.split()), y))
    sayac = defaultdict(list)
    for u in urunler:
        sayac[katla(u["baslik"])].append(u["kod"])
    uy, sonuc, ozet = [], [], defaultdict(int)
    if not any(u["marka"] for u in urunler):
        uy.append(uyari("Orta", "Marka sütunu yok", "", "Marka yazılmadığı için marka kontrolü ve önerisi yapılmadı."))
    for u in urunler:
        b, sorun = u["baslik"], []
        kel = katla(b).split()
        if len(b) < en_az:
            sorun.append("Kısa başlık")
        elif len(b) > en_cok:
            sorun.append("Uzun başlık")
        mi, ti = (_bul(kel, u["marka"]) if u["marka"] else None), (_bul(kel, u["tip"]) if u["tip"] else None)
        if u["marka"]:
            if mi is None:
                sorun.append("Marka yok")
            elif mi != 0:
                sorun.append("Marka başta değil")
        if u["tip"]:
            if ti is None:
                sorun.append("Ürün tipi yok")
            elif mi is not None and ti < mi:
                sorun.append("Ürün tipi markadan önce")
        for alan, ad in OZELLIK:
            if u[alan] and _bul(kel, u[alan], True) is None:
                sorun.append(f"{ad} yok")
        bulunan = [y for y in yasakli if _bul(kel, y) is not None]
        sorun += [f"Yasaklı ifade: {y}" for y in bulunan]
        harf = [c for c in b if c.isalpha()]
        if harf and sum(1 for c in harf if c.isupper()) * 2 > len(harf):
            sorun.append("Büyük harf")
        if OZEL.search(b):
            sorun.append("Özel karakter")
        tekrar = sorted({k for k in kel if len(k) >= 3 and not k.isdigit() and kel.count(k) > 1})
        sorun += [f"Tekrar eden kelime: {k}" for k in tekrar]
        if u["ham"] != b:
            sorun.append("Fazla boşluk")
        if len(sayac[katla(b)]) > 1:
            sorun.append("Aynı başlık")
        oneri = onerilen(u, bulunan, "Büyük harf" in sorun, (not u["marka"] or mi == 0) and (not u["tip"] or ti is not None))
        for s in sorun:
            ozet[s.split(":")[0]] += 1
        sonuc.append({**u, "uzunluk": len(b), "sorunlar": sorun, "oneri": oneri, "oneri_uzunluk": len(oneri), "esit": oneri == u["ham"],
                      "ayni": [k for k in sayac[katla(b)] if k != u["kod"]]})
    return {"urunler": sonuc, "ozet": dict(sorted(ozet.items(), key=lambda x: (-x[1], x[0]))), "uyarilar": uy}


def onerilen(u: dict, yasakli: list[str], buyuk: bool, sira_dogru: bool = False) -> str:
    """Sıra doğruysa başlığın kendi kelime sırası korunur; değilse Marka + Ürün tipi + kalan kelimeler."""
    kelimeler = OZEL.sub(" ", u["baslik"]).split()
    kat = [katla(k) for k in kelimeler]
    sil = set()
    for ifade in ([] if sira_dogru else [u["marka"], u["tip"]]) + yasakli:
        p = katla(ifade).split()
        if not p:
            continue
        for i in range(len(kat) - len(p) + 1):
            if kat[i:i + len(p)] == p:
                sil.update(range(i, i + len(p)))
    kalan, gorulen = [], set()
    for i, (k, kk) in enumerate(zip(kelimeler, kat)):
        if i in sil or not kk or (kk in gorulen and len(kk) >= 3 and not kk.isdigit()):
            continue
        gorulen.add(kk)
        kalan.append(tr_baslik(k) if buyuk and k.isupper() and len(k) > 1 else k)
    parca = ([] if sira_dogru else [x for x in (u["marka"], u["tip"]) if x]) + kalan
    metin_ = " ".join(parca)
    for alan, _ in OZELLIK:
        if u[alan] and _bul(katla(metin_).split(), u[alan], True) is None:
            parca.append(u[alan])
            metin_ = " ".join(parca)
    return metin_


def rapor_yaz(cikti: Path, r: dict, en_az: int, en_cok: int) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Başlık Kontrolü"
    _baslik(ws, ["Ürün Kodu", "Mevcut Başlık", "Uzunluk", "Sorun Sayısı", "Sorunlar", "Önerilen Başlık", "Öneri Uzunluğu", "Aynı Başlıklı Ürünler"], (14, 50, 8, 8, 40, 50, 8, 18), "C2")
    for s in r["urunler"]:
        n = len(s["sorunlar"])
        _satir(ws, [s["kod"], s["baslik"], s["uzunluk"], n, "; ".join(s["sorunlar"]) or "Uygun", "" if s["esit"] else s["oneri"], None if s["esit"] else s["oneri_uzunluk"], ", ".join(s["ayni"])],
               dolgu={4: KIRMIZI if n >= 3 else SARI if n else YESIL, 3: SARI if not en_az <= s["uzunluk"] <= en_cok else None,
                      7: SARI if not s["esit"] and not en_az <= s["oneri_uzunluk"] <= en_cok else None}, sar=(2, 5, 6))
    ws.append([])
    ws.append([f"Uzunluk {en_az}-{en_cok} karakter. Öneri otomatik üretilmiştir; anlam ve doğruluk kontrolü sizindir. Boş öneri: başlık zaten uygun."])
    ws = wb.create_sheet("Sorun Özeti")
    _baslik(ws, ["Sorun", "Ürün Sayısı"], (30, 10))
    for k, v in r["ozet"].items():
        _satir(ws, [k, v])
    uyari_sayfasi(wb, r["uyarilar"])
    kaydet(wb, cikti)


def calistir(yol: Path, cikti: Path, en_az: int = 30, en_cok: int = 100, yasakli=YASAKLI) -> dict:
    urunler, uy = oku(yol)
    r = kontrol(urunler, en_az, en_cok, yasakli)
    r["uyarilar"] = sirala(uy + r["uyarilar"])
    rapor_yaz(cikti, r, en_az, en_cok)
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Ürün başlıklarını uzunluk, marka-ürün tipi-özellik sırası ve yasaklı ifadeler açısından kontrol edip öneri sunar.")
    p.add_argument("--urunler", type=Path, help="Ürün Kodu, Başlık, Marka, Ürün Tipi, Renk, Materyal, Beden / Ebat, Model")
    p.add_argument("--min", type=int, default=30, help="En kısa başlık (karakter, varsayılan 30)")
    p.add_argument("--max", type=int, default=100, help="En uzun başlık (karakter, varsayılan 100)")
    p.add_argument("--yasakli", type=Path, help="Yasaklı ifadeler dosyası (ilk sütun, her satıra bir ifade); verilirse varsayılan listenin yerine geçer")
    p.add_argument("--ek-yasakli", default="", help="Listeye eklenecek ifadeler, noktalı virgülle: \"orijinal;garantili\"")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "urun_basligi_kontrolu.xlsx")
    a = p.parse_args(argv)
    yol = a.urunler or ORNEK / "urun_basliklari.csv"
    if a.min < 1 or a.max < a.min:
        print("[X] --min en az 1, --max --min'den küçük olamaz")
        return 2
    for y in (yol, a.yasakli):
        if y and not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    yasakli = (yasakli_oku(a.yasakli) if a.yasakli else list(YASAKLI)) + [x for x in a.ek_yasakli.split(";") if x.strip()]
    try:
        r = calistir(yol, a.cikti, a.min, a.max, yasakli)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    sorunlu = sum(1 for s in r["urunler"] if s["sorunlar"])
    print(f"[OK] {len(r['urunler'])} başlık · {sorunlu} başlıkta sorun · {len(r['urunler']) - sorunlu} uygun")
    for k, v in list(r["ozet"].items())[:8]:
        print(f"     {k}: {v}")
    for u in r["uyarilar"]:
        if u["onem"] in ("Yüksek", "Orta"):
            print(f"[!] {u['tur']}: {u['kim']} · {u['aciklama']}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
