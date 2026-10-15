"""
İhtarname Taslağı — Workers / Workless AI Agent
Hukuk › Avukat

1. Kod olay özetini satır satır numaralar; taraf bilgilerini (ad, adres, kimlik no) ayrı dosyadan okur ve olay metnindeki
   taraf adlarını {KESIDECI} / {MUHATAP} yer tutucusuyla değiştirir. Taraf bilgileri modele gönderilmez.
2. Model noter ihtarnamesinin konu, açıklamalar (her paragraf dayandığı satırlarla) ve sonuç-istem bölümlerini yazar.
3. Kod denetler:
   - Her paragrafın dayanak satırı var mı; paragraftaki tutar / tarih / sayılar dayandığı satırlarda yazıyor mu
   - Sonuç-istemde verilen süre (--sure gün) aynen yazılmış mı; başka sayı uydurulmuş mu
   - Kanun maddesi atfı olay metninde, talepte ya da --dayanak listesinde var mı (madde uydurulmasın)
   - Kira temerrüdü ihtarında (--tur kira-konut / kira-diger) TBK 315'teki asgari süre (30 / 10 gün) ve fesih ihtarı var mı
   - {KESIDECI} ve {MUHATAP} kullanılmış mı; yer tutucuya ek getirilmiş mi; hakaret sayılabilecek ifade var mı
   Sorun varsa modele geri gönderilir (en fazla --tur kez).
4. Kod taraf bilgilerini ve noterlik talep kalıbını ekleyip .txt taslağı ve kontrol Excel'ini yazar.

Kullanım:
    python agent.py                                                    # örnek (kurgusal) kira alacağı olayı
    python agent.py --olay olay.txt --taraflar taraflar.txt --talep "…" --sure 30 --tur kira-konut
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import belge
import llm

BURASI = Path(__file__).resolve().parent
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
AI_DOLGU = PatternFill("solid", fgColor="EDE7FF")
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
SARI = PatternFill("solid", fgColor="FFF4CE")
YESIL = PatternFill("solid", fgColor="E3F5E1")
UST = Alignment(vertical="top", wrap_text=True)
S = {"type": "string"}


def katla(s) -> str:
    """Karşılaştırma için: Türkçe harfleri sadeleştirir, küçük harfe çevirir, noktalamayı boşluk yapar."""
    return " ".join(re.sub(r"[^a-z0-9%]+", " ", str(s or "").replace("İ", "i").replace("I", "ı").translate(_TR).lower()).split())


def metinde_var(alinti: str, metin: str) -> bool:
    """Alıntı (boşluk, noktalama ve büyük/küçük harf farkı gözetilmeden) metinde geçiyor mu?"""
    a = katla(alinti)
    return bool(a) and a in katla(metin)


def sayilar(metin) -> set[str]:
    """Metindeki sayılar ('1.250,50' → '1250.5'; '08:30' → '08:30'; '15.10.2026' → '15.10.2026')."""
    s = str(metin or "")
    out = set(re.findall(r"\b\d{1,2}[.:]\d{2}(?:[.:]\d{2,4})?\b", s))
    for x in re.findall(r"\d+(?:[.,]\d+)*", s):
        out.add(sayi_norm(x))
    return out


def sayi_norm(x: str) -> str:
    x = str(x).strip()
    if re.fullmatch(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?", x):
        x = x.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?", x):
        x = x.replace(",", "")
    else:
        x = x.replace(",", ".")
    try:
        v = float(x)
    except ValueError:
        return x
    return sayi_yaz(v)


def sayi_yaz(v) -> str:
    """Üstel gösterim olmadan sade sayı: 19475.5 → '19475.5', 8450.0 → '8450'."""
    return f"{float(v):.6f}".rstrip("0").rstrip(".")


def tarih_coz(x) -> date | None:
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    s = str(x or "").strip()
    for f in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def zaman_coz(x) -> datetime | None:
    if isinstance(x, datetime):
        return x
    s = str(x or "").strip()
    for f in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    return None


def tablo_oku(yol: Path) -> list[list]:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: dosya bulunamadı.")
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, data_only=True, read_only=True)
        try:
            s = [list(r) for r in wb.active.iter_rows(values_only=True)]
        finally:
            wb.close()
    else:
        metin = yol.read_text(encoding="utf-8-sig")
        ilk = "\n".join(metin.splitlines()[:5])
        s = list(csv.reader(metin.splitlines(), delimiter=max(";\t,", key=ilk.count)))
    return [r for r in s if any(c not in (None, "") for c in r)]


def kayitlar(yol: Path, alanlar: dict[str, tuple], zorunlu: tuple = ()) -> list[dict]:
    """Başlık satırını eşanlamlılarla eşler; her satırı {alan: değer} sözlüğüne çevirir (boş hücre → '')."""
    s = tablo_oku(yol)
    if not s:
        raise llm.LLMHatasi(f"{yol.name}: tablo boş.")
    b = [katla(x) for x in s[0]]
    idx = {a: next((i for i, x in enumerate(b) if x in {katla(e) for e in es}), None) for a, es in alanlar.items()}
    eksik = [alanlar[a][0] for a in zorunlu if idx[a] is None]
    if eksik:
        raise llm.LLMHatasi(f"{yol.name}: gerekli sütun bulunamadı: {', '.join(eksik)}. Başlıklar: {s[0]}")
    out = []
    for r in s[1:]:
        k = {a: (r[i] if i is not None and i < len(r) and r[i] is not None else "") for a, i in idx.items()}
        k = {a: (v.strip() if isinstance(v, str) else v) for a, v in k.items()}
        out.append(k)
    return out


def baslik(ws, basliklar: list, genislik: tuple = ()) -> None:
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = UST
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w


def hizala(ws, ilk: int = 2) -> None:
    for satir in ws.iter_rows(min_row=ilk):
        for h in satir:
            h.alignment = UST


def boya(ws, sutunlar, dolgu) -> None:
    for c in sutunlar:
        ws.cell(ws.max_row, c).fill = dolgu


def duzeltme_dongusu(uret, denetle, tur: int) -> tuple[dict, list[str], int]:
    """uret(duzeltme, onceki) → taslak; denetle(taslak) → sorunlar. Sorun kalırsa en fazla `tur` kez yeniden yazdırır."""
    taslak = uret(None, None)
    kalan = denetle(taslak)
    yapilan = 0
    while kalan and yapilan < tur:
        yapilan += 1
        print(f"[i] Düzeltme turu {yapilan}: {len(kalan)} sorun modele geri gönderiliyor")
        taslak = uret(kalan, taslak)
        kalan = denetle(taslak)
    return taslak, kalan, yapilan


def duzeltme_blogu(duzeltme, onceki) -> str:
    if not duzeltme:
        return ""
    return ("\n<duzeltme>\nÖnceki taslak: " + json.dumps(onceki, ensure_ascii=False, default=str) + "\nKodun bulduğu sorunlar:\n"
            + "\n".join(f"- {s}" for s in duzeltme) + "\n</duzeltme>")


def ortam_hazirla() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass


def son_not(ws, metin: str, kalan: list[str], tur: int) -> None:
    ws.append([])
    ws.append(["Son kontrol: " + ("; ".join(kalan) if kalan else "Kod denetiminden sorunsuz geçti")])
    ws.cell(ws.max_row, 1).fill = KIRMIZI if kalan else YESIL
    ws.append([f"Düzeltme turu: {tur} · Model: {llm.kullanim_ozeti()} · {datetime.now():%d.%m.%Y %H:%M}"])
    ws.append([metin])
K = {"type": "array", "items": {"type": "integer"}}
SEMA = {
    "type": "object",
    "properties": {
        "konu": S,
        "aciklamalar": {"type": "array", "items": {"type": "object", "properties": {"metin": S, "kaynak_satirlar": K},
                                                   "required": ["metin", "kaynak_satirlar"], "additionalProperties": False}},
        "sonuc_ve_istem": S,
        "eksik_bilgiler": {"type": "array", "items": S},
        "ic_not": S,
    },
    "required": ["konu", "aciklamalar", "sonuc_ve_istem", "eksik_bilgiler", "ic_not"],
    "additionalProperties": False,
}
ASGARI = {"kira-konut": 30, "kira-diger": 10}  # TBK m.315: konut ve çatılı işyerinde en az 30, diğer kiralarda en az 10 gün
TARAF_ALANLARI = ("Noterlik", "Keşideci", "Keşideci Adres", "Keşideci Kimlik No", "Vekil", "Muhatap", "Muhatap Adres", "Muhatap Kimlik No")
MADDE = re.compile(r"(?:\b(?:TBK|TTK|TMK|HMK|İİK|IIK|TKHK)\s*(?:m\.|md\.)?\s*|\b(?:md|m)\.\s*)(\d{1,4})|\b(\d{1,4})\s*(?:\.|'?\s*[iıuü]nc[iıuü])?\s*madde", re.I)
HAKARET = re.compile(r"(dolandırıcı|hırsız|sahtekâr|sahtekar|namussuz|rezil|yalancı|pişman edeceğ|gününüzü göreceksiniz)", re.I)
EK = re.compile(r"\{(KESIDECI|MUHATAP)\}['’]?[a-zçğıöşü]", re.I)
NOTER_TALEBI = ("Sayın Noter; işbu ihtarnamenin üç nüsha olarak düzenlenmesini, bir nüshasının muhataba tebliğini, bir nüshasının "
                "dairenizde saklanmasını ve tebliğ şerhli bir nüshasının tarafıma verilmesini saygıyla rica ederim.")


def satirlar(metin: str) -> list[str]:
    return [s.strip() for s in metin.splitlines() if s.strip()]


def taraflari_oku(yol: Path) -> dict:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: taraf bilgileri dosyası bulunamadı.")
    t = {}
    for s in belge.metin_oku(yol).splitlines():
        a, _, d = s.partition(":")
        alan = next((x for x in TARAF_ALANLARI if katla(x) == katla(a)), None)
        if alan and d.strip():
            t[alan] = d.strip()
    eksik = [x for x in ("Keşideci", "Muhatap") if x not in t]
    if eksik:
        raise llm.LLMHatasi(f"{yol.name}: zorunlu satır yok: {', '.join(eksik)} (ör. 'Muhatap: Ad Soyad').")
    return t


def gizle(metin: str, t: dict) -> str:
    """Olay metnindeki taraf adlarını yer tutucuyla değiştirir (büyük/küçük harf farkı gözetmeden)."""
    for alan, yer in (("Keşideci", "{KESIDECI}"), ("Muhatap", "{MUHATAP}"), ("Vekil", "{VEKIL}")):
        if t.get(alan):
            metin = re.sub(re.escape(t[alan]), yer, metin, flags=re.I)
    return llm.maskele(metin)


def denetle(tas: dict, l: list[str], talep: str, sure: int, tur_: str, dayanak: str) -> list[str]:
    n, sorun = len(l), []
    izinli = "\n".join(l) + " " + talep + " " + dayanak + (" TBK 315" if tur_ in ASGARI else "")
    izinli_madde = {x for m in MADDE.finditer(izinli) for x in m.groups() if x} | {x for x in re.findall(r"\d+", dayanak)}
    if not tas["aciklamalar"]:
        sorun.append("Açıklamalar: paragraf yok")
    for i, p in enumerate(tas["aciklamalar"], 1):
        ks = p["kaynak_satirlar"]
        if not ks:
            sorun.append(f"Paragraf {i}: dayanak satır yok")
        if [k for k in ks if not 1 <= k <= n]:
            sorun.append(f"Paragraf {i}: olayda olmayan satır {[k for k in ks if not 1 <= k <= n]}")
        kaynak = " ".join(l[k - 1] for k in ks if 1 <= k <= n)
        maddeler = {x for m in MADDE.finditer(p["metin"]) for x in m.groups() if x}
        fazla = sorted(sayilar(MADDE.sub(" ", p["metin"])) - sayilar(kaynak))
        if fazla:
            sorun.append(f"Paragraf {i}: dayandığı satırlarda olmayan sayı / tarih: {', '.join(fazla)}")
        for m in sorted(maddeler - izinli_madde):
            sorun.append(f"Paragraf {i}: madde {m} atfı olayda, talepte ya da dayanak listesinde yok (madde uydurulmasın)")
    si = tas["sonuc_ve_istem"]
    if not re.search(rf"\b{sure}\s*(\(.*?\)\s*)?gün", si, re.I):
        sorun.append(f"Sonuç ve istem: verilen süre '{sure} gün' yazılmamış")
    fazla = sorted(sayilar(MADDE.sub(" ", si)) - sayilar("\n".join(l) + " " + talep) - {str(sure)})
    if fazla:
        sorun.append(f"Sonuç ve istem: olayda ve talepte olmayan sayı: {', '.join(fazla)}")
    for m in sorted({x for m in MADDE.finditer(si + " " + tas["konu"]) for x in m.groups() if x} - izinli_madde):
        sorun.append(f"Sonuç ve istem: madde {m} atfı olayda, talepte ya da dayanak listesinde yok")
    if tur_ in ASGARI and not re.search(r"fesh", si, re.I):
        sorun.append("Sonuç ve istem: kira temerrüdü ihtarında süre içinde ödenmezse sözleşmenin feshedileceği yazılmamış (TBK 315)")
    tum = " ".join([tas["konu"], si] + [p["metin"] for p in tas["aciklamalar"]])
    for y in ("{KESIDECI}", "{MUHATAP}"):
        if y not in tum:
            sorun.append(f"Genel: {y} yer tutucusu kullanılmamış")
    if EK.search(tum):
        sorun.append(f"Genel: yer tutucuya ek getirilmiş ('{EK.search(tum).group(0)}'); cümleyi ek gerektirmeyecek biçimde kurun")
    if HAKARET.search(tum):
        sorun.append(f"Genel: hakaret / tehdit sayılabilecek ifade ('{HAKARET.search(tum).group(0)}')")
    return sorun


def uret(l: list[str], talep: str, sure: int, tur_: str, dayanak: str, duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    olay = "\n".join(f"{i}. {s}" for i, s in enumerate(l, 1))
    mesaj = (f"<olay>\n{olay}\n</olay>\n<talep>{talep}</talep>\n<sure_gun>{sure}</sure_gun>\n<ihtar_turu>{tur_}</ihtar_turu>\n"
             f"<dayanak_maddeler>{dayanak or '(verilmedi)'}</dayanak_maddeler>" + duzeltme_blogu(duzeltme, onceki))
    return llm.json_iste(sistem, mesaj, SEMA)


def doldur(metin: str, t: dict) -> str:
    return (metin.replace("{KESIDECI}", t.get("Keşideci", "")).replace("{MUHATAP}", t.get("Muhatap", ""))
            .replace("{VEKIL}", t.get("Vekil", "")))


def metin_yaz(tas: dict, t: dict, tarih: date) -> str:
    s = [t.get("Noterlik", "…………… NOTERLİĞİNE").upper(), ""]
    s.append(f"KEŞİDECİ      : {t['Keşideci']}" + (f" (T.C. Kimlik / Vergi No: {t['Keşideci Kimlik No']})" if t.get("Keşideci Kimlik No") else ""))
    if t.get("Keşideci Adres"):
        s.append(f"                {t['Keşideci Adres']}")
    if t.get("Vekil"):
        s.append(f"VEKİLİ        : {t['Vekil']}")
    s.append(f"MUHATAP       : {t['Muhatap']}" + (f" (T.C. Kimlik / Vergi No: {t['Muhatap Kimlik No']})" if t.get("Muhatap Kimlik No") else ""))
    s.append(f"                {t.get('Muhatap Adres', '[ADRES EKSİK — tebligat için zorunlu]')}")
    s += [f"KONU          : {doldur(tas['konu'], t)}", "", "AÇIKLAMALAR   :", ""]
    s += [f"{i}. {doldur(p['metin'], t)}" for i, p in enumerate(tas["aciklamalar"], 1)]
    s += ["", "SONUÇ VE İSTEM:", "", doldur(tas["sonuc_ve_istem"], t), "", NOTER_TALEBI, "", f"Tarih: {tarih:%d.%m.%Y}",
          f"Keşideci{' Vekili' if t.get('Vekil') else ''}: {t.get('Vekil') or t['Keşideci']}", "İmza: ", "",
          "TASLAKTIR — avukat kontrolü ve imzası olmadan notere verilmez."]
    return "\n".join(s)


def rapor_yaz(cikti: Path, tas: dict, l: list[str], t: dict, kalan, tur, metin: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Paragraflar"
    baslik(ws, ["No", "Paragraf", "Dayanak Satırlar", "Dayanak Metni"], (5, 80, 12, 70))
    for i, p in enumerate(tas["aciklamalar"], 1):
        ws.append([i, doldur(p["metin"], t), ", ".join(map(str, p["kaynak_satirlar"])),
                   "\n".join(f"{k}. {l[k - 1]}" for k in p["kaynak_satirlar"] if 1 <= k <= len(l))])
        boya(ws, [2], AI_DOLGU)
    ws.append(["SONUÇ", doldur(tas["sonuc_ve_istem"], t)])
    boya(ws, [2], AI_DOLGU)
    son_not(ws, "Önemli: Bu bir taslaktır. Hukuki nitelendirme, süre ve talepler avukat tarafından kontrol edilmelidir.", kalan, tur)
    hizala(ws)
    e = wb.create_sheet("Eksik Bilgiler")
    baslik(e, ["Eksik / Teyit Edilecek Bilgi"], (90,))
    for x in tas["eksik_bilgiler"] + ([] if t.get("Muhatap Adres") else ["Muhatap adresi (tebligat için zorunlu)"]):
        e.append([x])
        boya(e, [1], SARI)
    e.append(["İç not: " + tas["ic_not"]])
    hizala(e)
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)
    cikti.with_suffix(".txt").write_text(metin, encoding="utf-8")


def calistir(olay_yolu: Path, taraf_yolu: Path, talep: str, sure: int, cikti: Path, tur_: str = "genel", dayanak: str = "",
             tur: int = 2, evet: bool = False, tarih: date | None = None) -> dict:
    if not olay_yolu.exists():
        raise llm.LLMHatasi(f"{olay_yolu}: olay özeti bulunamadı.")
    if tur_ in ASGARI and sure < ASGARI[tur_]:
        raise llm.LLMHatasi(f"--sure {sure}: {tur_} ihtarında TBK m.315'e göre süre en az {ASGARI[tur_]} gün olmalıdır.")
    t = taraflari_oku(taraf_yolu)
    l = satirlar(gizle(belge.metin_oku(olay_yolu), t))
    if not l:
        raise llm.LLMHatasi(f"{olay_yolu.name}: olay özeti boş.")
    print(f"[OK] {len(l)} olay satırı · süre {sure} gün · tür {tur_}")
    llm.onay_al(f"{len(l)} satırlık olay özeti (taraf adları {{KESIDECI}} / {{MUHATAP}} ile değiştirildi; adres ve kimlik no gönderilmez) ve talep gönderilecek.", evet)
    tas, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(l, talep, sure, tur_, dayanak, d, o), lambda x: denetle(x, l, talep, sure, tur_, dayanak), tur)
    metin = metin_yaz(tas, t, tarih or date.today())
    rapor_yaz(cikti, tas, l, t, kalan, yapilan, metin)
    return {"taslak": tas, "kalan": kalan, "tur": yapilan, "metin": metin, "satirlar": l}


def main(argv: list[str] | None = None) -> int:
    ortam_hazirla()
    o = BURASI / "ornek_veri"
    p = argparse.ArgumentParser(description="Olay özeti ve talepten noter ihtarnamesi taslağı hazırlar (avukat onayı gerekir).")
    p.add_argument("--olay", type=Path, default=o / "olay.txt", help="Olay özeti (.txt/.docx/.pdf); her satır bir olgu")
    p.add_argument("--taraflar", type=Path, default=o / "taraflar.txt", help="'Keşideci: …', 'Muhatap: …', 'Muhatap Adres: …' satırları")
    p.add_argument("--talep", default="Ödenmeyen kira bedellerinin ödenmesi; ödenmezse kira sözleşmesinin feshi ve tahliye",
                   help="Muhataptan istenen")
    p.add_argument("--sure", type=int, default=30, help="Muhataba verilecek süre (gün, varsayılan 30)")
    p.add_argument("--tur-ihtar", dest="tur_", choices=["genel", "kira-konut", "kira-diger"], default="kira-konut",
                   help="kira-konut / kira-diger: TBK m.315 asgari süre denetimi (30 / 10 gün)")
    p.add_argument("--dayanak", default="", help="Atıf yapılabilecek maddeler, ör. \"TBK 315, TBK 117\" (verilmezse yalnız olayda geçenler)")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "ihtarname_taslagi.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    try:
        s = calistir(a.olay, a.taraflar, a.talep, a.sure, a.cikti, a.tur_, a.dayanak, a.tur, a.evet)
    except (llm.LLMHatasi, belge.BelgeHatasi) as h:
        print(f"[X] {h}")
        return 1
    print(f"[OK] {len(s['taslak']['aciklamalar'])} paragraflık ihtarname taslağı")
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    print(f"[OK] Taslak: {a.cikti.with_suffix('.txt').resolve()}")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
