"""
Kaizen Önerisi Ön İnceleme — Workers / Workless AI Agent
Endüstri Mühendisliği › Endüstri Mühendisi (Süreç Geliştirme Uzmanı da kullanır)

Çalışanların serbest metinle yazdığı iyileştirme önerilerini değerlendirmeye hazır kayda çevirir.
1. Kod öneri listesini okur; öneren adları sağlayıcıya gönderilmez, telefon / e-posta / kimlik numarası maskelenir.
2. Model her öneri için kısa konu başlığı, kategori (İSG, kalite, verimlilik, maliyet, enerji, 5S / ergonomi), metinde
   yazan kazanç bileşenleri (zaman, malzeme, doğrudan TL), uygulama maliyeti, uygulanabilirlik önerisi (1-5, gerekçeli),
   benzer / aynı öneriler ve önerenden istenecek eksik bilgileri çıkarır.
3. Kod çıkarımı denetler: her sayı öneri metninde geçiyor mu; birim ve periyot tanımlı mı; benzer öneri numarası listede
   mi. Sorun varsa yalnız o öneri modele geri gönderilir (en fazla --tur kez).
4. Kod yıllık kazancı sizin parametrelerinizle hesaplar:
   zaman = dakika × tekrar × kişi × yıllık periyot × saat maliyeti ÷ 60;
   malzeme = miktar × birim fiyat × yıllık periyot; doğrudan TL = tutar × yıllık periyot.
   İSG önerileri kazançtan bağımsız işaretlenir; kazanç hesaplanamayanlar "bilgi iste" olarak ayrılır.
Çıktının ilk sayfası "Kaizen Önerisi Değerlendirme" kod bloğunun girdi biçimindedir.

Kullanım:
    python agent.py                                                     # örnek 6 öneri
    python agent.py --oneriler oneriler.xlsx --saat-maliyeti 420 --gun 255
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import llm

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
AI_DOLGU = PatternFill("solid", fgColor="EDE7FF")
SARI = PatternFill("solid", fgColor="FFF4CE")
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
YESIL = PatternFill("solid", fgColor="E3F5E1")
UST = Alignment(vertical="top", wrap_text=True)
S = {"type": "string"}
KATEGORI = ["İSG", "Kalite", "Verimlilik", "Maliyet", "Enerji", "5S / Ergonomi", "Diğer"]
BIRIM = ["dakika", "saat", "kg", "adet", "litre", "kWh", "TL"]
PERIYOT = ["gün", "hafta", "ay", "yıl", "vardiya"]


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9%]+", " ", str(s or "").replace("İ", "i").replace("I", "ı").translate(_TR).lower()).split())


def para(x) -> Decimal | None:
    if x in (None, ""):
        return None
    if isinstance(x, (int, float, Decimal)):
        return Decimal(str(x))
    s = re.sub(r"[^\d,.\-]", "", str(x))
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[1]) == 3):
        s = s.replace(".", "")
    try:
        return Decimal(s) if s not in ("", "-", ".") else None
    except InvalidOperation:
        return None


SOZ = {"bir": 1, "iki": 2, "uc": 3, "dort": 4, "bes": 5, "alti": 6, "yedi": 7, "sekiz": 8, "dokuz": 9, "on": 10, "yarim": Decimal("0.5")}


def sayilar(metin: str) -> set[Decimal]:
    """Metindeki rakamlar ve 'iki kişi', 'yarım saat' gibi sözle yazılmış küçük sayılar."""
    out = {v for x in re.findall(r"\d+(?:[.,]\d+)*", str(metin)) if (v := para(x)) is not None}
    out |= {Decimal(SOZ[k]) for k in katla(metin).split() if k in SOZ}
    return out | {Decimal(1)}            # tekrar / kişi alanlarında "1" (yazılmamış) her zaman kabul


def tl(x: Decimal) -> str:
    return f"{x:,.0f}".replace(",", ".")


# ----------------------------------------------------------------------------
# Girdi
# ----------------------------------------------------------------------------

SUTUN = {"no": ("oneri no", "no", "oneri numarasi"), "oneren": ("oneren", "ad soyad", "calisan"), "bolum": ("bolum", "birim", "departman"),
         "metin": ("oneri metni", "oneri", "aciklama", "metin")}


def onerileri_oku(yol: Path) -> list[dict]:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: öneri dosyası bulunamadı.")
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, read_only=True, data_only=True)
        s = [list(r) for r in wb.active.iter_rows(values_only=True)]
        wb.close()
    else:
        metin = yol.read_text(encoding="utf-8-sig")
        s = list(csv.reader(metin.splitlines(), delimiter=max(";\t,", key=metin.split("\n")[0].count)))
    s = [r for r in s if any(c not in (None, "") for c in r)]
    if not s:
        raise llm.LLMHatasi(f"{yol.name}: dosya boş.")
    b = [katla(c) for c in s[0]]
    idx = {a: next((i for i, x in enumerate(b) if x in es), None) for a, es in SUTUN.items()}
    if idx["metin"] is None:
        raise llm.LLMHatasi(f"{yol.name}: 'Öneri Metni' sütunu gerekli. Başlıklar: {s[0]}")
    out = []
    for i, r in enumerate(s[1:], 1):
        al = lambda a: "" if idx[a] is None or idx[a] >= len(r) or r[idx[a]] is None else str(r[idx[a]]).strip()  # noqa: E731
        if al("metin"):
            out.append({"no": al("no") or f"K-{i:03d}", "oneren": al("oneren"), "bolum": al("bolum"), "metin": al("metin")})
    nolar = [o["no"] for o in out]
    if len(set(nolar)) != len(nolar):
        raise llm.LLMHatasi(f"{yol.name}: öneri numaraları tekil değil.")
    return out


def maskele(o: dict) -> str:
    m = o["metin"]
    for parca in sorted([o["oneren"], *o["oneren"].split()], key=len, reverse=True):
        if len(parca) > 2:
            m = re.sub(rf"\b{re.escape(parca)}\b", "[ÖNEREN]", m, flags=re.I)
    return llm.maskele(m)


# ----------------------------------------------------------------------------
# Model
# ----------------------------------------------------------------------------

BILESEN = {"type": "object", "properties": {
    "tur": {"type": "string", "enum": ["Zaman", "Malzeme", "Doğrudan TL"]}, "miktar": S, "birim": {"type": "string", "enum": BIRIM},
    "tekrar": S, "kisi": S, "periyot": {"type": "string", "enum": PERIYOT}, "birim_fiyat": S, "aciklama": S},
    "required": ["tur", "miktar", "birim", "tekrar", "kisi", "periyot", "birim_fiyat", "aciklama"], "additionalProperties": False}
SEMA = {"type": "object", "properties": {
    "konu": S, "kategori": {"type": "string", "enum": KATEGORI}, "kazanc": {"type": "array", "items": BILESEN},
    "uygulama_maliyeti": S, "uygulanabilirlik": {"type": "integer"}, "uygulanabilirlik_gerekcesi": S,
    "benzer": {"type": "array", "items": S}, "eksik_bilgiler": {"type": "array", "items": S}},
    "required": ["konu", "kategori", "kazanc", "uygulama_maliyeti", "uygulanabilirlik", "uygulanabilirlik_gerekcesi", "benzer", "eksik_bilgiler"],
    "additionalProperties": False}


def denetle(o: dict, c: dict, nolar: set[str]) -> list[str]:
    on, say, sorun = o["no"], sayilar(o["metin"]), []

    def sayi(ad, v, bos_olabilir=True):
        if not v.strip():
            if not bos_olabilir:
                sorun.append(f"{on}: {ad} boş")
            return
        d = para(v)
        if d is None or d not in say:
            sorun.append(f"{on}: {ad} '{v}' öneri metninde yazmıyor (tahmin etmeyin; yoksa eksik bilgilere yazın)")

    for j, k in enumerate(c["kazanc"], 1):
        ad = f"Kazanç {j}"
        sayi(f"{ad} miktar", k["miktar"], False)
        sayi(f"{ad} tekrar", k["tekrar"])
        sayi(f"{ad} kişi", k["kisi"])
        sayi(f"{ad} birim fiyat", k["birim_fiyat"])
        if k["tur"] == "Zaman" and k["birim"] not in ("dakika", "saat"):
            sorun.append(f"{on}: {ad}: zaman kazancının birimi dakika ya da saat olmalı")
        if k["tur"] == "Doğrudan TL" and k["birim"] != "TL":
            sorun.append(f"{on}: {ad}: doğrudan TL kazancının birimi TL olmalı")
        if k["tur"] == "Malzeme" and not k["birim_fiyat"].strip():
            sorun.append(f"{on}: {ad}: malzeme kazancı için birim fiyat metinde yoksa bileşeni yazmayın, eksik bilgilere ekleyin")
    sayi("Uygulama maliyeti", c["uygulama_maliyeti"])
    if not 1 <= c["uygulanabilirlik"] <= 5:
        sorun.append(f"{on}: uygulanabilirlik 1-5 arasında olmalı")
    if not c["uygulanabilirlik_gerekcesi"].strip():
        sorun.append(f"{on}: uygulanabilirlik gerekçesi boş")
    yok = [b for b in c["benzer"] if b not in nolar or b == on]
    if yok:
        sorun.append(f"{on}: benzer öneri {', '.join(yok)} listede yok ya da kendisi")
    if not c["konu"].strip() or len(c["konu"]) > 90:
        sorun.append(f"{on}: konu boş ya da 90 karakterden uzun")
    if "[ÖNEREN]" in c["konu"]:
        sorun.append(f"{on}: konuda öneren yer tutucusu var")
    return sorun


def yillik_kazanc(c: dict, saat: Decimal, gun: int, vardiya: int) -> tuple[Decimal | None, list[str]]:
    """Kod: bileşenlerden yıllık TL ve hesap satırları. Hesaplanamazsa None."""
    yil = {"gün": Decimal(gun), "hafta": Decimal(gun) / 5, "ay": Decimal(12), "yıl": Decimal(1), "vardiya": Decimal(gun * vardiya)}
    toplam, satir = Decimal(0), []
    for k in c["kazanc"]:
        m = para(k["miktar"])
        if m is None:
            continue
        t, ki, f = para(k["tekrar"]) or Decimal(1), para(k["kisi"]) or Decimal(1), para(k["birim_fiyat"])
        p = yil[k["periyot"]]
        if k["tur"] == "Zaman":
            dk = m * (60 if k["birim"] == "saat" else 1)
            v = dk * t * ki * p * saat / 60
            satir.append(f"{m:g} {k['birim']} × {t:g} tekrar × {ki:g} kişi × {p:g} ({k['periyot']}/yıl) × {tl(saat)} TL/saat = {tl(v)} TL")
        elif k["tur"] == "Malzeme" and f is not None:
            v = m * t * f * p
            satir.append(f"{m:g} {k['birim']} × {t:g} × {f:g} TL × {p:g} ({k['periyot']}/yıl) = {tl(v)} TL")
        elif k["tur"] == "Doğrudan TL":
            v = m * t * p
            satir.append(f"{tl(m)} TL × {t:g} × {p:g} ({k['periyot']}/yıl) = {tl(v)} TL")
        else:
            continue
        toplam += v
    return (toplam.quantize(Decimal(1), rounding=ROUND_HALF_UP) if satir else None), satir


def uret(oneriler: list[dict], duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    liste = "\n".join(f"{o['no']} | {o['bolum']} | {o['maskeli'][:140]}" for o in oneriler)
    sonuc = dict(onceki or {})
    hedef = oneriler if duzeltme is None else [o for o in oneriler if any(s.startswith(o["no"] + ":") for s in duzeltme)]
    for o in hedef:
        mesaj = f"<tum_oneriler>\n{liste}\n</tum_oneriler>\n<oneri no=\"{o['no']}\" bolum=\"{o['bolum']}\">\n{o['maskeli']}\n</oneri>"
        d = [s.split(": ", 1)[1] for s in (duzeltme or []) if s.startswith(o["no"] + ":")]
        if d:
            mesaj += ("\n<duzeltme>\nÖnceki inceleme: " + json.dumps(sonuc.get(o["no"]), ensure_ascii=False) + "\nKodun bulduğu sorunlar:\n"
                      + "\n".join(f"- {s}" for s in d) + "\n</duzeltme>")
        sonuc[o["no"]] = llm.json_iste(sistem, mesaj, SEMA)
    return sonuc


def duzeltme_dongusu(uret_f, denetle_f, tur: int) -> tuple[dict, list[str], int]:
    taslak = uret_f(None, None)
    kalan = denetle_f(taslak)
    yapilan = 0
    while kalan and yapilan < tur:
        yapilan += 1
        print(f"[i] Düzeltme turu {yapilan}: {len(kalan)} sorun modele geri gönderiliyor")
        taslak = uret_f(kalan, taslak)
        kalan = denetle_f(taslak)
    return taslak, kalan, yapilan


def degerlendir(oneriler: list[dict], sonuc: dict, saat: Decimal, gun: int, vardiya: int) -> list[dict]:
    out = []
    for o in oneriler:
        c = sonuc[o["no"]]
        kz, hesap = yillik_kazanc(c, saat, gun, vardiya)
        durum = []
        if c["kategori"] == "İSG":
            durum.append("İSG önerisi: kazanç puanından bağımsız, İSG kurulunda değerlendirin")
        if kz is None and c["kategori"] != "İSG":
            durum.append("Kazanç hesaplanamadı: önerenden bilgi isteyin")
        if kz is not None and para(c["uygulama_maliyeti"]) is None:
            durum.append("Uygulama maliyeti yazılmamış: değerlendirme kod bloğu boş maliyeti 0 sayar ve geri ödeme puanı olduğundan yüksek çıkar; "
                         "maliyeti ekleyin")
        if c["benzer"]:
            durum.append(f"Benzer öneri: {', '.join(c['benzer'])} (birlikte değerlendirin; ilk önereni kaydedin)")
        out.append({"o": o, "c": c, "kazanc": kz, "hesap": hesap, "maliyet": para(c["uygulama_maliyeti"]), "durum": durum})
    return out


# ----------------------------------------------------------------------------
# Çıktı
# ----------------------------------------------------------------------------

def _baslik(ws, basliklar, genislik) -> None:
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font, h.alignment = BASLIK_DOLGU, BASLIK_YAZI, UST
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def rapor_yaz(cikti: Path, sonuclar: list[dict], kalan: list[str], tur: int, parametre: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Öneriler"
    _baslik(ws, ["Öneri No", "Öneren", "Bölüm", "Konu", "Yıllık Kazanç", "Uygulama Maliyeti", "Uygulanabilirlik", "Kategori", "Kazanç Hesabı",
                 "Uygulanabilirlik Gerekçesi", "Benzer Öneriler", "Eksik Bilgiler", "Durum", "Öneri Metni", "Son Kontrol"],
            (9, 16, 12, 40, 12, 12, 9, 12, 60, 40, 12, 45, 45, 70, 30))
    for s in sonuclar:
        o, c = s["o"], s["c"]
        sor = [x.split(": ", 1)[1] for x in kalan if x.startswith(o["no"] + ":")]
        ws.append([o["no"], o["oneren"], o["bolum"], c["konu"], None if s["kazanc"] is None else float(s["kazanc"]),
                   None if s["maliyet"] is None else float(s["maliyet"]), c["uygulanabilirlik"], c["kategori"], "\n".join(s["hesap"]),
                   c["uygulanabilirlik_gerekcesi"], ", ".join(c["benzer"]), "\n".join(f"• {x}" for x in c["eksik_bilgiler"]), "\n".join(s["durum"]),
                   o["metin"], "; ".join(sor) or "Geçti"])
        n = ws.max_row
        ws.cell(n, 5).number_format = ws.cell(n, 6).number_format = "#,##0"
        for j in (4, 7, 8, 10, 11, 12):
            ws.cell(n, j).fill = AI_DOLGU
        if s["durum"]:
            ws.cell(n, 13).fill = SARI
        ws.cell(n, 15).fill = KIRMIZI if sor else YESIL
        for h in ws[n]:
            h.alignment = UST
    k = wb.create_sheet("Kontrol")
    k.append([f"Parametreler: {parametre}"])
    k.append([f"Düzeltme turu: {tur} · Model: {llm.kullanim_ozeti()} · {datetime.now():%d.%m.%Y %H:%M}"])
    k.append(["Son kontrol: " + ("; ".join(kalan) if kalan else "Kod denetiminden sorunsuz geçti")])
    k.cell(3, 1).fill = KIRMIZI if kalan else YESIL
    k.append(["Önemli: Konu, kategori, uygulanabilirlik ve benzer öneri yapay zekâ önerisidir (mor). Yıllık kazanç, önerinin metnindeki "
              "rakamlardan ve yukarıdaki parametrelerden kodla hesaplanmıştır; önerenle ve ilgili bölümle doğrulayın. Uygulanabilirlik "
              "sütununu gözden geçirip dosyayı 'Kaizen Önerisi Değerlendirme' kod bloğuna verin."])
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(girdi: Path, cikti: Path, saat: Decimal = Decimal(300), gun: int = 250, vardiya: int = 2, tur: int = 2, evet: bool = False) -> dict:
    oneriler = onerileri_oku(girdi)
    for o in oneriler:
        o["maskeli"] = maskele(o)
    nolar = {o["no"] for o in oneriler}
    print(f"[OK] {len(oneriler)} öneri")
    llm.onay_al(f"{len(oneriler)} öneri metni ve bölüm adları gönderilecek. Öneren adları gönderilmez; telefon, e-posta, TCKN maskeli.", evet)
    sonuc, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(oneriler, d, o),
                                             lambda x: [s for o in oneriler for s in denetle(o, x[o["no"]], nolar)], tur)
    sonuclar = degerlendir(oneriler, sonuc, saat, gun, vardiya)
    rapor_yaz(cikti, sonuclar, kalan, yapilan, f"saat maliyeti {tl(saat)} TL · yılda {gun} iş günü · günde {vardiya} vardiya · hafta = {gun}/5 gün")
    return {"sonuclar": sonuclar, "kalan": kalan, "tur": yapilan}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    p = argparse.ArgumentParser(description="Serbest metin iyileştirme önerilerini kazanç hesabı ve eksik bilgileriyle değerlendirmeye hazırlar.")
    p.add_argument("--oneriler", type=Path, default=ORNEK / "oneriler.csv", help="Öneriler (.xlsx/.csv): Öneri No, Öneren, Bölüm, Öneri Metni")
    p.add_argument("--saat-maliyeti", default="300", help="İşçilik saat maliyeti TL (varsayılan 300; kendi değerinizi verin)")
    p.add_argument("--gun", type=int, default=250, help="Yıllık çalışma günü (varsayılan 250)")
    p.add_argument("--vardiya", type=int, default=2, help="Günlük vardiya sayısı (periyot 'vardiya' için; varsayılan 2)")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "kaizen_on_inceleme.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    saat = para(a.saat_maliyeti)
    if saat is None or saat <= 0 or a.gun <= 0 or a.vardiya <= 0:
        print("[X] --saat-maliyeti, --gun ve --vardiya pozitif olmalı")
        return 2
    try:
        s = calistir(a.oneriler, a.cikti, saat, a.gun, a.vardiya, a.tur, a.evet)
    except llm.LLMHatasi as h:
        print(f"[X] {h}")
        return 1
    for x in s["sonuclar"]:
        kz = "hesaplanamadı" if x["kazanc"] is None else f"{tl(x['kazanc'])} TL/yıl"
        print(f"[OK] {x['o']['no']} · {x['c']['kategori']} · {x['c']['konu'][:50]} · {kz}")
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
