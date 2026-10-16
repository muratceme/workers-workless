"""
İkinci El Araç İlan Metni — Workers / Workless AI Agent
Otomotiv › Satış › İkinci El Satış Danışmanı

Araç kartı ve ekspertiz notlarından ilan başlığı, açıklama, donanım listesi ve ekspertiz özeti yazar; ekspertizin
eksiksiz ve doğru yansıtıldığını kod denetler.
1. Kod araç kartlarını ("İkinci El Araç Fiyat Değerleme" kod bloğunun araç dosyası + Renk, İlan Fiyatı) ve ekspertiz
   notlarını ("Araç: Stok-001" blokları; Boyalı / Lokal boyalı / Değişen / Tramer / Bakım / Donanım / Not satırları)
   okur. Telefon, e-posta, TCKN maskelenir.
2. Model başlık, açıklama, donanım ve ekspertiz özetini yazar.
3. Kod denetler: başlıkta yıl, metinde km var mı; her sayı kartta ya da notlarda var mı; her boyalı / lokal boyalı /
   değişen parça ve tramer tutarı ekspertiz özetinde mi; ekspertizle çelişen ifade ("boyasız", "değişensiz",
   "kazasız", "tramersiz", "hatasız" …) var mı; donanım notta yazıyor mu; başlık uzunluğu. Sorun varsa yalnız o araç
   modele geri gönderilir.
4. Kod --ek-satir ile verilen sabit satırı (ör. yetki belgesi no, iletişim) açıklamanın sonuna ekler.

Kullanım:
    python agent.py                                                         # örnek 2 araç
    python agent.py --araclar araclar.xlsx --ekspertiz ekspertiz.docx --baslik-sinir 60 --ek-satir "Yetki Belgesi No: …"
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import belge
import llm

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
AI_DOLGU = PatternFill("solid", fgColor="EDE7FF")
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
YESIL = PatternFill("solid", fgColor="E3F5E1")
UST = Alignment(vertical="top", wrap_text=True)
S = {"type": "string"}
YOK = {"", "yok", "-", "hayir", "temiz", "0"}
HER_ZAMAN = re.compile(r"\b(hatasiz|kusursuz|sifir ayarinda|sifirdan farksiz|orjinal|orijinal)\b")
BOYA = re.compile(r"\b(boyasiz|boyasi yok|boya yok|boyali parca yok)\w*")
DEGISEN = re.compile(r"\b(degisensiz|degiseni yok|degisen yok|degisen parca yok)\w*")
HASAR = re.compile(r"\b(kazasiz|tramersiz|hasarsiz|hasar kaydi yok|tramer kaydi yok|tramer yok)\w*")


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9%]+", " ", str(s or "").replace("İ", "i").replace("I", "ı").translate(_TR).lower()).split())


def sayi(x) -> Decimal | None:
    s = re.sub(r"[^\d,.\-]", "", str(x if x is not None else ""))
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1 or re.fullmatch(r"\d{1,3}\.\d{3}", s):
        s = s.replace(".", "")
    try:
        return Decimal(s) if s not in ("", "-", ".") else None
    except InvalidOperation:
        return None


def sayilar(metin: str) -> set[Decimal]:
    return {v for x in re.findall(r"\d+(?:[.,]\d+)*", str(metin)) if (v := sayi(x)) is not None}


def parcalar(x: str) -> list[str]:
    return [] if katla(x) in YOK else [p.strip() for p in re.split(r"[,;/]| ve ", x) if katla(p) not in YOK]


# ----------------------------------------------------------------------------
# Girdi
# ----------------------------------------------------------------------------

def araclari_oku(yol: Path) -> dict[str, dict]:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: araç dosyası bulunamadı.")
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, read_only=True, data_only=True)
        s = [list(r) for r in wb.active.iter_rows(values_only=True)]
        wb.close()
    else:
        metin = yol.read_text(encoding="utf-8-sig")
        s = list(csv.reader(metin.splitlines(), delimiter=max(";\t,", key=metin.split("\n")[0].count)))
    s = [r for r in s if any(c not in (None, "") for c in r)]
    bas = [katla(c) for c in s[0]]
    out = {}
    for r in s[1:]:
        d = {k: ("" if v is None else str(v).strip()) for k, v in zip(bas, list(r) + [None] * (len(bas) - len(r)))}
        no = d.get("arac") or d.get("stok no") or d.get("plaka")
        if no:
            out[no] = {"no": no, "marka": d.get("marka", ""), "model": d.get("model", ""), "yil": d.get("yil", ""), "km": d.get("km", ""), "yakit": d.get("yakit", ""),
                       "vites": d.get("vites", ""), "paket": d.get("paket", ""), "renk": d.get("renk", ""), "fiyat": d.get("ilan fiyati") or d.get("fiyat", "")}
    if not out:
        raise llm.LLMHatasi(f"{yol.name}: 'Araç' sütunu olan araç listesi gerekli.")
    return out


def ekspertiz_oku(yol: Path) -> dict[str, dict]:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: ekspertiz notları bulunamadı.")
    out, son = {}, None
    for s in belge.metin_oku(yol).replace("\r", "").splitlines():
        m = re.match(r"^\s*Araç\s*:\s*(\S+)", s, re.I)
        if m:
            son = out.setdefault(m[1], {"boyali": [], "lokal": [], "degisen": [], "tramer": None, "bakim": "", "donanim": [], "not": ""})
            continue
        if not son or not s.strip():
            continue
        k, _, v = s.partition(":")
        kk = katla(k)
        v = llm.maskele(v.strip())
        if kk == "boyali":
            son["boyali"] = parcalar(v)
        elif kk == "lokal boyali":
            son["lokal"] = parcalar(v)
        elif kk == "degisen":
            son["degisen"] = parcalar(v)
        elif kk in ("tramer", "hasar kaydi"):
            son["tramer"] = None if katla(v) in YOK else sayi(v)
        elif kk == "bakim":
            son["bakim"] = v
        elif kk == "donanim":
            son["donanim"] = [x.strip() for x in v.split(",") if x.strip()]
        else:
            son["not"] = (son["not"] + " " + (v if kk == "not" else llm.maskele(s.strip()))).strip()
    return out


def ekspertiz_metni(e: dict) -> str:
    return "\n".join([f"Boyalı: {', '.join(e['boyali']) or 'yok'}", f"Lokal boyalı: {', '.join(e['lokal']) or 'yok'}", f"Değişen: {', '.join(e['degisen']) or 'yok'}",
                      f"Tramer: {e['tramer']:,.0f} TL".replace(",", ".") if e["tramer"] else "Tramer: yok", f"Bakım: {e['bakim'] or '-'}",
                      f"Donanım: {', '.join(e['donanim']) or '-'}", f"Not: {e['not'] or '-'}"])


# ----------------------------------------------------------------------------
# Model
# ----------------------------------------------------------------------------

SEMA = {"type": "object", "properties": {"baslik": S, "aciklama": S, "donanim": {"type": "array", "items": S}, "ekspertiz_ozeti": S},
        "required": ["baslik", "aciklama", "donanim", "ekspertiz_ozeti"], "additionalProperties": False}


def denetle(a: dict, e: dict, c: dict, baslik_sinir: int) -> list[str]:
    no, sorun = a["no"], []
    hepsi = " ".join([c["baslik"], c["aciklama"], c["ekspertiz_ozeti"], *c["donanim"]])
    kh, ko = katla(hepsi), katla(c["ekspertiz_ozeti"])
    if len(c["baslik"]) > baslik_sinir:
        sorun.append(f"{no}: başlık {len(c['baslik'])} karakter (sınır {baslik_sinir})")
    if a["yil"] and a["yil"] not in c["baslik"]:
        sorun.append(f"{no}: başlıkta model yılı ({a['yil']}) yok")
    if a["km"] and sayi(a["km"]) not in sayilar(c["baslik"] + " " + c["aciklama"]):
        sorun.append(f"{no}: metinde km ({a['km']}) yok ya da farklı yazılmış")
    izinli = sayilar(" ".join(a.values()) + " " + ekspertiz_metni(e))
    fazla = sorted(sayilar(hepsi) - izinli, key=float)
    if fazla:
        sorun.append(f"{no}: kartta ve notlarda olmayan sayı: {', '.join(f'{x:,.0f}'.replace(',', '.') for x in fazla)}")
    for tur, liste in (("boyalı", e["boyali"]), ("lokal boyalı", e["lokal"]), ("değişen", e["degisen"])):
        eksik = [p for p in liste if katla(p) not in ko]
        if eksik:
            sorun.append(f"{no}: {tur} parça ekspertiz özetinde yok: {', '.join(eksik)}")
    if e["tramer"]:
        if e["tramer"] not in sayilar(c["ekspertiz_ozeti"]) or not re.search(r"\b(tramer|hasar kaydi)", ko):
            sorun.append(f"{no}: tramer (hasar kaydı) tutarı {e['tramer']:,.0f} TL ekspertiz özetinde yok".replace(",", "."))
    for kural, kosul, ad in ((HER_ZAMAN, True, "kanıtlanamayan ifade"), (BOYA, bool(e["boyali"] or e["lokal"]), "boyalı parça varken"),
                             (DEGISEN, bool(e["degisen"]), "değişen parça varken"), (HASAR, bool(e["tramer"] or e["degisen"]), "tramer / değişen varken")):
        if kosul and (g := kural.search(kh)):
            sorun.append(f"{no}: {ad} '{g[0]}' ifadesi kullanılmış")
    kn = katla(" ".join(e["donanim"]) + " " + e["not"] + " " + a["paket"])
    uydurma = [d for d in c["donanim"] if katla(d) not in kn]
    if uydurma:
        sorun.append(f"{no}: notlarda olmayan donanım: {', '.join(uydurma)}")
    if not c["ekspertiz_ozeti"].strip():
        sorun.append(f"{no}: ekspertiz özeti boş")
    return sorun


def uret(araclar: list[dict], eksp: dict, duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    sonuc = dict(onceki or {})
    hedef = araclar if duzeltme is None else [a for a in araclar if any(s.startswith(a["no"] + ":") for s in duzeltme)]
    for a in hedef:
        kart = "\n".join(f"{k}: {v}" for k, v in (("Marka", a["marka"]), ("Model", a["model"]), ("Yıl", a["yil"]), ("Km", a["km"]), ("Yakıt", a["yakit"]),
                                                  ("Vites", a["vites"]), ("Paket", a["paket"]), ("Renk", a["renk"]), ("Fiyat", f"{a['fiyat']} TL" if a["fiyat"] else "")) if v)
        mesaj = f"<arac no=\"{a['no']}\">\n{kart}\n</arac>\n<ekspertiz>\n{ekspertiz_metni(eksp[a['no']])}\n</ekspertiz>"
        d = [s.split(": ", 1)[1] for s in (duzeltme or []) if s.startswith(a["no"] + ":")]
        if d:
            mesaj += ("\n<duzeltme>\nÖnceki cevap: " + json.dumps(sonuc.get(a["no"]), ensure_ascii=False) + "\nKodun bulduğu sorunlar:\n"
                      + "\n".join(f"- {s}" for s in d) + "\n</duzeltme>")
        sonuc[a["no"]] = llm.json_iste(sistem, mesaj, SEMA)
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


# ----------------------------------------------------------------------------
# Çıktı
# ----------------------------------------------------------------------------

def ilan_metni(a: dict, c: dict, ek_satir: str) -> str:
    l = [c["baslik"], "", c["aciklama"].strip(), ""]
    if c["donanim"]:
        l += ["Donanım:"] + [f"- {d}" for d in c["donanim"]] + [""]
    l += ["Ekspertiz:", c["ekspertiz_ozeti"].strip()]
    if ek_satir:
        l += ["", ek_satir]
    return "\n".join(l)


def rapor_yaz(cikti: Path, araclar: list[dict], sonuc: dict, kalan: list[str], tur: int, ek_satir: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "İlanlar"
    ws.append(["Araç", "Marka / Model", "Yıl", "Km", "Fiyat", "Başlık", "Başlık Uzunluğu", "Açıklama", "Donanım", "Ekspertiz Özeti", "Son Kontrol"])
    for h in ws[1]:
        h.fill, h.font, h.alignment = BASLIK_DOLGU, BASLIK_YAZI, UST
    for a in araclar:
        c = sonuc[a["no"]]
        sor = "; ".join(x.split(": ", 1)[1] for x in kalan if x.startswith(a["no"] + ":"))
        ws.append([a["no"], f"{a['marka']} {a['model']}", a["yil"], a["km"], a["fiyat"], c["baslik"], len(c["baslik"]), c["aciklama"], "\n".join(c["donanim"]),
                   c["ekspertiz_ozeti"], sor or "Geçti"])
        r = ws.max_row
        for j in (6, 8, 9, 10):
            ws.cell(r, j).fill = AI_DOLGU
        ws.cell(r, 11).fill = YESIL if not sor else KIRMIZI
        for h in ws[r]:
            h.alignment = UST
    for j, w in enumerate((12, 22, 7, 10, 12, 45, 10, 80, 35, 60, 35), 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    k = wb.create_sheet("Kontrol")
    k.append([f"Düzeltme turu: {tur} · Model: {llm.kullanim_ozeti()} · {datetime.now():%d.%m.%Y %H:%M}"])
    k.append(["Son kontrol: " + ("; ".join(kalan) if kalan else "Kod denetiminden sorunsuz geçti")])
    k.cell(2, 1).fill = KIRMIZI if kalan else YESIL
    k.append(["Mor hücreler modelin metinleridir. Yayınlamadan önce ekspertiz raporu ve tramer sorgusuyla karşılaştırın; ilan platformunun ve "
              "yürürlükteki mevzuatın ilanda istediği bilgileri ekleyin. İlan metinleri aynı adlı .txt dosyasındadır."])
    k.column_dimensions["A"].width = 140
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)
    txt = cikti.with_suffix(".txt")
    txt.write_text("\n\n==========\n\n".join(f"[{a['no']}]\n" + ilan_metni(a, sonuc[a["no"]], ek_satir) for a in araclar) + "\n", encoding="utf-8")
    return txt


def calistir(a_yol: Path, e_yol: Path, cikti: Path, baslik_sinir: int = 60, ek_satir: str = "", secili: list[str] | None = None, tur: int = 2, evet: bool = False) -> dict:
    kartlar, eksp = araclari_oku(a_yol), ekspertiz_oku(e_yol)
    araclar = [kartlar[n] for n in (secili or kartlar) if n in kartlar]
    eksik = [a["no"] for a in araclar if a["no"] not in eksp]
    if eksik:
        raise llm.LLMHatasi(f"Ekspertiz notu olmayan araç: {', '.join(eksik)}. Notu olmayan araç için ilan yazılmaz.")
    if not araclar:
        raise llm.LLMHatasi("İlan yazılacak araç yok.")
    print(f"[OK] {len(araclar)} araç · ekspertiz notu {len(eksp)}")
    llm.onay_al(f"{len(araclar)} aracın kartı ve ekspertiz notları gönderilecek (plaka ve şasi no gönderilmez; telefon, e-posta, TCKN maskeli).", evet)
    sonuc, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(araclar, eksp, d, o), lambda x: [s for a in araclar for s in denetle(a, eksp[a["no"]], x[a["no"]], baslik_sinir)], tur)
    txt = rapor_yaz(cikti, araclar, sonuc, kalan, yapilan, ek_satir)
    return {"araclar": araclar, "sonuc": sonuc, "kalan": kalan, "tur": yapilan, "txt": txt}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    p = argparse.ArgumentParser(description="Araç kartı ve ekspertiz notlarından ikinci el ilan metni yazar; ekspertizin eksiksiz yansıtıldığını kod denetler.")
    p.add_argument("--araclar", type=Path, default=ORNEK / "araclar.csv", help="Araç; Marka; Model; Yıl; Km; Yakıt; Vites; Paket; Renk; İlan Fiyatı")
    p.add_argument("--ekspertiz", type=Path, default=ORNEK / "ekspertiz_notlari.txt", help="'Araç: No' blokları: Boyalı, Lokal boyalı, Değişen, Tramer, Bakım, Donanım, Not")
    p.add_argument("--arac", nargs="*", default=[], metavar="NO", help="Yalnız bu araçlar (varsayılan hepsi)")
    p.add_argument("--baslik-sinir", type=int, default=60, help="İlan başlığı en fazla karakter; platformunuzun sınırını girin (varsayılan 60)")
    p.add_argument("--ek-satir", default="", help="Her ilanın sonuna eklenecek sabit satır (ör. yetki belgesi no, iletişim)")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "ilan_metinleri.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    try:
        s = calistir(a.araclar, a.ekspertiz, a.cikti, a.baslik_sinir, a.ek_satir, a.arac, a.tur, a.evet)
    except (llm.LLMHatasi, belge.BelgeHatasi) as h:
        print(f"[X] {h}")
        return 1
    for x in s["araclar"]:
        print(f"[OK] {x['no']} · {s['sonuc'][x['no']]['baslik']}")
    for x in s["kalan"]:
        print(f"[!] Kalan: {x}")
    print(f"[OK] Rapor: {a.cikti.resolve()} (+ {s['txt'].name})")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
