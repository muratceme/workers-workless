"""
Toplu Ödeme Dosyası Kontrolü — Gerisi Bende kod bloğu
Bankacılık Operasyonları › Ödeme Sistemleri Operasyon Uzmanı

Müşterilerin yüklediği toplu ödeme (maaş, tedarikçi) dosyalarını işleme almadan önce kontrol eder:
  - IBAN: boşluk / küçük harf temizlenir; TR IBAN 26 karakter, yalnız rakam, 11. karakter (rezerv) "0";
    kontrol basamağı ISO 13616 mod 97 (ülke kodu + kontrol basamağı sona alınır, harfler A=10 … Z=35, kalan 1 olmalı).
    Yabancı IBAN mod 97 ile doğrulanır ve ayrıca işaretlenir.
  - Tutar: sıfırdan büyük, en çok 2 ondalık; --ust-sinir verilirse tek ödeme sınırı.
  - Alıcı adı boş / çok kısa; TCKN ve VKN verilmişse kontrol basamakları.
  - Mükerrer: aynı IBAN + aynı tutar (olası çift ödeme); aynı IBAN farklı alıcı adı; aynı kimlik no farklı IBAN.
  - Gönderen hesabına ödeme (--gonderen-iban).
  - Toplam kontrolü: --beklenen-adet ve --beklenen-toplam (müşteri talimatındaki kontrol toplamı) ile karşılaştırma;
    dosyanın sonundaki "TOPLAM" satırı ödeme sayılmaz, tutarı ayrıca karşılaştırılır.
  - Önceki dosya (--onceki): aynı kimlik no / alıcı adına ait IBAN değişmişse (dolandırıcılık riski: teyit),
    tutar --degisim-esik yüzdesinden fazla değişmişse, önceki dosyanın birebir tekrarıysa.
Kod yalnız biçim ve tutarlılığı kontrol eder; hesabın açık olup olmadığını veya IBAN'ın alıcıya ait olduğunu bilemez.
İnternete bağlanmaz.

Kullanım:
    python main.py                                                 # örnek: 40 satırlık maaş dosyası, önceki ayla
    python main.py --dosya odemeler.xlsx --beklenen-adet 120 --beklenen-toplam 4.250.000,00 --onceki gecen_ay.xlsx
"""
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


SUTUNLAR = {"ad": ("alici adi", "alici", "ad soyad", "alici adi soyadi", "unvan", "alici unvani"), "iban": ("iban", "alici iban", "iban no", "hesap no"),
            "tutar": ("tutar", "odeme tutari", "net odenecek", "net tutar"), "kimlik": ("tckn vkn", "tckn", "vkn", "tc kimlik no", "kimlik no", "vergi no"),
            "aciklama": ("aciklama", "odeme aciklamasi"), "sira": ("sira", "sira no", "no")}
# IBAN'ın 5-9. karakterleri (banka kodu). Liste bilgi amaçlıdır ve eksik olabilir; güncel liste için TCMB / BKM kayıtlarına bakın.
BANKALAR = {"00001": "TCMB", "00010": "Ziraat Bankası", "00012": "Halkbank", "00015": "VakıfBank", "00032": "TEB", "00046": "Akbank", "00059": "Şekerbank",
            "00062": "Garanti BBVA", "00064": "İş Bankası", "00067": "Yapı Kredi", "00099": "ING", "00103": "Fibabanka", "00111": "QNB", "00123": "HSBC",
            "00124": "Alternatif Bank", "00134": "DenizBank", "00135": "Anadolubank", "00143": "Aktif Bank", "00146": "Odeabank", "00203": "Albaraka Türk",
            "00205": "Kuveyt Türk", "00206": "Türkiye Finans", "00209": "Ziraat Katılım", "00210": "Vakıf Katılım", "00211": "Emlak Katılım"}


def iban_temizle(x) -> str:
    return re.sub(r"[\s\-.]", "", metin(x)).upper()


def iban_mod97(iban: str) -> bool:
    if not re.fullmatch(r"[A-Z]{2}\d{2}[A-Z0-9]{10,30}", iban):
        return False
    return int("".join(str(int(c, 36)) for c in iban[4:] + iban[:4])) % 97 == 1


def iban_kontrol(x) -> tuple[str, str | None]:
    """(temiz IBAN, hata) — hata None ise geçerli."""
    i = iban_temizle(x)
    if not i:
        return i, "IBAN boş"
    if not re.fullmatch(r"[A-Z]{2}[A-Z0-9]+", i):
        return i, "IBAN ülke koduyla başlamalı ve yalnız harf / rakam içermeli"
    if i.startswith("TR"):
        if len(i) != 26:
            return i, f"TR IBAN 26 karakter olmalı ({len(i)} karakter)"
        if not i[2:].isdigit():
            return i, "TR IBAN'da ülke kodundan sonra yalnız rakam olur"
        if i[10] != "0":
            return i, "TR IBAN'ın 11. karakteri (rezerv alan) 0 olmalı"
    if not iban_mod97(i):
        return i, "IBAN kontrol basamağı hatalı (mod 97)"
    return i, None


def tckn_gecerli(s: str) -> bool:
    if not re.fullmatch(r"[1-9]\d{10}", s):
        return False
    d = [int(c) for c in s]
    return ((sum(d[0:9:2]) * 7 - sum(d[1:8:2])) % 10 == d[9]) and (sum(d[:10]) % 10 == d[10])


def vkn_gecerli(s: str) -> bool:
    if not re.fullmatch(r"\d{10}", s):
        return False
    t = 0
    for i in range(9):
        v = (int(s[i]) + 9 - i) % 10
        if v:
            v = (v * 2 ** (9 - i)) % 9 or 9
        t += v
    return (10 - t % 10) % 10 == int(s[9])


def ad_katla(x) -> str:
    return katla(x)


def tl(x) -> str:
    return (f"{x:,.2f}" if isinstance(x, Decimal) else f"{x:,}").replace(",", "X").replace(".", ",").replace("X", ".")


def oku(yol: Path) -> tuple[list[dict], Decimal | None]:
    """Satırları ve varsa dosya sonundaki TOPLAM satırının tutarını döndürür."""
    satirlar, dip = [], None
    for r in kayitlar(yol, SUTUNLAR, ("iban", "tutar")):
        if katla(r["ad"]).startswith("toplam") or katla(r["iban"]).startswith("toplam") or katla(r["sira"]).startswith("toplam"):
            dip = para(r["tutar"])
            continue
        if not (metin(r["ad"]) or metin(r["iban"]) or metin(r["tutar"])):
            continue
        satirlar.append({"satir": r["_satir"], "ad": metin(r["ad"]), "iban_ham": metin(r["iban"]), "tutar_ham": r["tutar"], "kimlik": re.sub(r"\D", "", metin(r["kimlik"])),
                         "kimlik_ham": metin(r["kimlik"]), "aciklama": metin(r["aciklama"])})
    return satirlar, dip


def kontrol(satirlar: list[dict], dip: Decimal | None = None, beklenen_adet: int | None = None, beklenen_toplam: Decimal | None = None, gonderen: str = "",
            ust_sinir: Decimal | None = None, onceki: list[dict] | None = None, degisim_esik: Decimal = Decimal(25)) -> dict:
    genel = []

    def b(s, onem, tur, aciklama):
        s["bulgular"].append((onem, tur, aciklama))

    gonderen = iban_temizle(gonderen)
    for s in satirlar:
        s["bulgular"] = []
        s["iban"], hata = iban_kontrol(s["iban_ham"])
        s["iban_hatali"] = bool(hata)
        s["banka"] = BANKALAR.get(s["iban"][4:9], "Diğer / tanımsız") if s["iban"].startswith("TR") and len(s["iban"]) == 26 else ""
        if hata:
            b(s, "Yüksek", "IBAN hatalı", hata)
        elif not s["iban"].startswith("TR"):
            b(s, "Orta", "Yabancı IBAN", f"Ülke kodu {s['iban'][:2]}. Yurt içi toplu ödeme kanalından gönderilemeyebilir.")
        if gonderen and s["iban"] == gonderen:
            b(s, "Yüksek", "Gönderen hesabına ödeme", "Alıcı IBAN'ı gönderen hesapla aynı.")
        t = para(s["tutar_ham"])
        s["tutar"] = t
        if t is None:
            b(s, "Yüksek", "Tutar okunamadı", f"'{metin(s['tutar_ham'])}'")
        elif t <= 0:
            b(s, "Yüksek", "Tutar sıfır veya eksi", tl(t))
        elif t != yuvarla(t):
            b(s, "Yüksek", "Tutar 2 ondalıktan fazla", f"{t:f}")
        elif ust_sinir is not None and t > ust_sinir:
            b(s, "Orta", "Tek ödeme sınırı aşıldı", f"{tl(t)} > {tl(ust_sinir)}")
        if len(ad_katla(s["ad"])) < 3:
            b(s, "Yüksek", "Alıcı adı eksik", f"'{s['ad']}'")
        if s["kimlik_ham"]:
            k = s["kimlik"]
            if len(k) == 11 and not tckn_gecerli(k):
                b(s, "Orta", "TCKN hatalı", "Kontrol basamakları tutmuyor.")
            elif len(k) == 10 and not vkn_gecerli(k):
                b(s, "Orta", "VKN hatalı", "Kontrol basamağı tutmuyor.")
            elif len(k) not in (10, 11):
                b(s, "Orta", "Kimlik no uzunluğu hatalı", f"{len(k)} hane; TCKN 11, VKN 10 hanedir.")
    # --- mükerrerler
    gecerli = [s for s in satirlar if s["tutar"] is not None and s["iban"]]
    cift = defaultdict(list)
    for s in gecerli:
        cift[(s["iban"], s["tutar"])].append(s)
    for l in cift.values():
        if len(l) > 1:
            for s in l[1:]:
                b(s, "Yüksek", "Mükerrer ödeme", f"Aynı IBAN ve tutar satır {l[0]['satir']} ile aynı. Çift ödeme olabilir.")
    for alan, tur, mesaj in (("iban", "Aynı IBAN, farklı alıcı", "ad"), ("kimlik", "Aynı kimlik no, farklı IBAN", "iban")):
        grup = defaultdict(list)
        for s in gecerli:
            if s[alan]:
                grup[s[alan]].append(s)
        for l in grup.values():
            deger = {ad_katla(s[mesaj]) if mesaj == "ad" else s[mesaj] for s in l}
            if len(deger) > 1:
                for s in l:
                    b(s, "Orta", tur, "Satırlar: " + ", ".join(str(x["satir"]) for x in l))
    # --- önceki dosya
    if onceki is not None:
        for o in onceki:
            o["iban"], o["tutar"] = iban_temizle(o["iban_ham"]), para(o["tutar_ham"])
        eski_kimlik = {o["kimlik"]: o for o in onceki if o["kimlik"]}
        eski_ad = {ad_katla(o["ad"]): o for o in onceki if ad_katla(o["ad"])}
        ayni, bulunan = 0, set()
        for s in gecerli:
            o = eski_kimlik.get(s["kimlik"]) or eski_ad.get(ad_katla(s["ad"]))             # önce kimlik no, yoksa ad
            if not o:
                b(s, "Bilgi", "Yeni alıcı", "Önceki dosyada yok.")
                continue
            bulunan.add(id(o))
            if o["iban"] != s["iban"] and not s["iban_hatali"]:
                b(s, "Yüksek", "IBAN değişmiş", f"Önceki dosyada {o['iban']}. Değişikliği alıcıyla / müşteriyle bağımsız bir kanaldan teyit edin.")
            if o["tutar"] and s["tutar"] and o["tutar"] != s["tutar"]:
                d = (s["tutar"] - o["tutar"]) / o["tutar"] * 100
                if abs(d) > degisim_esik:
                    b(s, "Orta", "Tutar belirgin değişmiş", f"Önceki {tl(o['tutar'])} → {tl(s['tutar'])} (%{yuvarla(d, '0.1')}).")
            ayni += o["iban"] == s["iban"] and o["tutar"] == s["tutar"]
        cikan = [o for o in onceki if id(o) not in bulunan]
        if cikan:
            genel.append(uyari("Bilgi", "Önceki dosyada olup bu dosyada olmayan alıcı", f"{len(cikan)} alıcı", ", ".join(o["ad"] for o in cikan[:15]) + (" …" if len(cikan) > 15 else "")))
        if gecerli and ayni == len(gecerli) == len(onceki):
            genel.append(uyari("Yüksek", "Dosya önceki dosyanın birebir aynısı", "Tüm satırlar", "Tüm alıcı, IBAN ve tutarlar önceki dosyayla aynı. Yanlış dosya yüklenmiş olabilir."))
    # --- toplamlar
    toplam = sum((s["tutar"] for s in satirlar if s["tutar"] is not None and s["tutar"] > 0), Decimal(0))
    for ad, beklenen, bulunan in (("Adet", beklenen_adet, len(satirlar)), ("Toplam tutar", beklenen_toplam, toplam), ("Dosyadaki TOPLAM satırı", dip, toplam)):
        if beklenen is not None and beklenen != bulunan:
            genel.append(uyari("Yüksek", "Kontrol toplamı tutmuyor", ad, f"Beklenen {tl(beklenen)} · dosyada {tl(bulunan)} · fark {tl(bulunan - beklenen)}"))
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    for s in satirlar:
        s["bulgular"].sort(key=lambda x: sira[x[0]])
        s["durum"] = "Hatalı" if any(x[0] == "Yüksek" for x in s["bulgular"]) else "Kontrol edin" if any(x[0] == "Orta" for x in s["bulgular"]) else "Temiz"
    banka = defaultdict(lambda: [0, Decimal(0)])
    for s in satirlar:
        if s["banka"] and s["tutar"]:
            banka[s["banka"]][0] += 1
            banka[s["banka"]][1] += s["tutar"]
    temiz = [s for s in satirlar if s["durum"] != "Hatalı"]
    return {"satirlar": satirlar, "genel": sirala(genel), "toplam": toplam, "adet": len(satirlar), "temiz_adet": len(temiz),
            "temiz_toplam": sum((s["tutar"] for s in temiz if s["tutar"]), Decimal(0)), "banka": dict(banka), "dip": dip,
            "beklenen": (beklenen_adet, beklenen_toplam), "sonuc": "İşleme alınamaz" if any(s["durum"] == "Hatalı" for s in satirlar) or any(u["onem"] == "Yüksek" for u in genel)
            else "Kontrol sonrası işlenebilir" if any(s["durum"] == "Kontrol edin" for s in satirlar) else "Biçim kontrolünden geçti"}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

def iban_goster(i: str) -> str:
    return " ".join(i[j:j + 4] for j in range(0, len(i), 4))


def rapor_yaz(cikti: Path, r: dict) -> None:
    wb = Workbook()
    oz = wb.active
    oz.title = "Özet"
    _baslik(oz, ["Bilgi", "Değer", "Açıklama"], (34, 24, 80))
    n = _satir(oz, ["Sonuç", r["sonuc"], ""], dolgu={2: {"İşleme alınamaz": KIRMIZI, "Kontrol sonrası işlenebilir": SARI}.get(r["sonuc"], YESIL)})
    oz.cell(n, 2).font = Font(bold=True)
    say = {d: sum(s["durum"] == d for s in r["satirlar"]) for d in ("Hatalı", "Kontrol edin", "Temiz")}
    for k, v, a in (("Satır sayısı", r["adet"], f"beklenen {r['beklenen'][0]}" if r["beklenen"][0] is not None else ""),
                    ("Toplam tutar", r["toplam"], f"beklenen {tl(r['beklenen'][1])}" if r["beklenen"][1] is not None else ""),
                    ("Dosyadaki TOPLAM satırı", r["dip"], ""), ("Hatalı satır", say["Hatalı"], "Düzeltilmeden işlenemez"), ("Kontrol edilecek satır", say["Kontrol edin"], ""),
                    ("Temiz satır", say["Temiz"], ""), ("Hatalılar hariç toplam", r["temiz_toplam"], f"{r['temiz_adet']} satır")):
        _satir(oz, [k, v, a], para_s=(2,) if isinstance(v, Decimal) else ())
    oz.append([])
    _baslik(oz, ["Banka (IBAN banka koduna göre)", "Adet", "Tutar"], dondur=None)
    for ad, (adet, tutar) in sorted(r["banka"].items(), key=lambda x: -x[1][1]):
        _satir(oz, [ad, adet, tutar], para_s=(3,))
    oz.append([])
    oz.append(["Kod biçim ve tutarlılığı kontrol eder; IBAN'ın alıcıya ait olduğunu veya hesabın açık olduğunu doğrulamaz."])

    ks = wb.create_sheet("Satır Kontrolü")
    _baslik(ks, ["Dosya Satırı", "Durum", "Alıcı", "IBAN", "Banka", "Tutar", "Kimlik No", "Açıklama", "Bulgular"], (8, 13, 30, 34, 18, 15, 13, 26, 100))
    for s in sorted(r["satirlar"], key=lambda s: ({"Hatalı": 0, "Kontrol edin": 1, "Temiz": 2}[s["durum"]], s["satir"])):
        _satir(ks, [s["satir"], s["durum"], s["ad"], iban_goster(s["iban"]) or s["iban_ham"], s["banka"], s["tutar"] if s["tutar"] is not None else metin(s["tutar_ham"]),
                    s["kimlik_ham"], s["aciklama"], " | ".join(f"{t}: {a}" for _, t, a in s["bulgular"])], para_s=(6,),
               dolgu={2: {"Hatalı": KIRMIZI, "Kontrol edin": SARI, "Temiz": YESIL}[s["durum"]]}, sar=(9,))
    ks.auto_filter.ref = f"A1:I{ks.max_row}"

    bl = wb.create_sheet("Bulgular")
    _baslik(bl, ["Önem", "Bulgu", "Dosya Satırı", "Alıcı", "Açıklama", "Müşteriye Bildirildi"], (9, 30, 9, 30, 100, 12))
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    for u in r["genel"]:
        _satir(bl, [u["onem"], u["tur"], "", u["kim"], u["aciklama"], ""], dolgu={1: RENK[u["onem"]], 6: SARI}, sar=(5,))
    for onem, tur, acik, s in sorted(((o, t, a, s) for s in r["satirlar"] for o, t, a in s["bulgular"]), key=lambda x: (sira[x[0]], x[1], x[3]["satir"])):
        _satir(bl, [onem, tur, s["satir"], s["ad"], acik, ""], dolgu={1: RENK[onem], 6: SARI}, sar=(5,))
    kaydet(wb, cikti)


def calistir(yol: Path, cikti: Path, onceki_yol: Path | None = None, **kw) -> dict:
    satirlar, dip = oku(yol)
    if not satirlar:
        raise ValueError(f"{yol.name}: ödeme satırı bulunamadı")
    r = kontrol(satirlar, dip, onceki=oku(onceki_yol)[0] if onceki_yol else None, **kw)
    rapor_yaz(cikti, r)
    return r


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Toplu ödeme dosyasında format, IBAN, mükerrer satır ve toplam kontrolü yapar.")
    p.add_argument("--dosya", type=Path, default=ORNEK / "odemeler.csv", help="Alıcı Adı, IBAN, Tutar, TCKN/VKN, Açıklama")
    p.add_argument("--onceki", type=Path, help="Önceki dönemin ödeme dosyası (IBAN değişikliği ve tutar karşılaştırması için)")
    p.add_argument("--beklenen-adet", type=int, help="Müşteri talimatındaki satır sayısı")
    p.add_argument("--beklenen-toplam", help="Müşteri talimatındaki toplam tutar (ör. 1.250.000,00)")
    p.add_argument("--gonderen-iban", default="", help="Ödemelerin çıkacağı hesabın IBAN'ı")
    p.add_argument("--ust-sinir", help="Tek ödeme için üst sınır")
    p.add_argument("--degisim-esik", type=Decimal, default=Decimal(25), help="Önceki dosyaya göre tutar değişimi uyarı eşiği %% (varsayılan 25)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "toplu_odeme_kontrolu.xlsx")
    a = p.parse_args(argv)
    ornek = a.dosya == ORNEK / "odemeler.csv"
    onceki = a.onceki or (ORNEK / "onceki_ay.csv" if ornek else None)
    for y in (a.dosya, onceki):
        if y and not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    bt, us = para(a.beklenen_toplam), para(a.ust_sinir)
    if (a.beklenen_toplam and bt is None) or (a.ust_sinir and us is None):
        print("[X] --beklenen-toplam ve --ust-sinir sayı olmalı (ör. 1.250.000,00)")
        return 2
    if ornek and a.beklenen_adet is None and bt is None:
        a.beklenen_adet, bt = 40, Decimal("1702740.00")       # örnek talimat: 40 satır
    try:
        r = calistir(a.dosya, a.cikti, onceki, beklenen_adet=a.beklenen_adet, beklenen_toplam=bt, gonderen=a.gonderen_iban, ust_sinir=us, degisim_esik=a.degisim_esik)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    hatali = [s for s in r["satirlar"] if s["durum"] == "Hatalı"]
    print(f"[OK] {r['adet']} satır · toplam {tl(r['toplam'])} · hatalı {len(hatali)} · sonuç: {r['sonuc']}")
    for u in r["genel"]:
        if u["onem"] == "Yüksek":
            print(f"[!] {u['tur']} · {u['kim']}: {u['aciklama']}")
    for s in hatali[:20]:
        print(f"[!] Satır {s['satir']} {s['ad']}: " + "; ".join(t for o, t, _ in s["bulgular"] if o == "Yüksek"))
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
