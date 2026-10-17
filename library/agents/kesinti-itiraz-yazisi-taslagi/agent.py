"""
Kesinti İtiraz Yazısı Taslağı — Gerisi Bende AI Agent
Sağlık › Anlaşmalı Kurumlar Uzmanı (SGK Faturalama Uzmanı da kullanır)

Hastane faturalarına sigorta şirketi / kurum tarafından yapılan kesintiler için, kurumun kendi belgelerine ve anlaşma /
mevzuat metnine dayanan itiraz yazısı taslağı hazırlar.
1. Kod kesinti listesini ("Kesinti ve İtiraz Takibi" kod bloğunun girdi biçimi), fatura no ile adlandırılmış destek
   belgelerini (belgeler/F-201.txt …) ve dayanak metnini (anlaşma maddeleri, SUT maddeleri — siz verirsiniz) okur.
   Hasta adları (--gizle), TCKN, telefon, e-posta ve IBAN sağlayıcıya gönderilmez.
2. Model her kesinti için itiraz önerisini (Evet / Hayır / Belge eksik), gerekçesini, dayandığı maddeleri ve belgeleri,
   eksik belgeleri ve itiraz yazısı metnini hazırlar.
3. Kod taslağı denetler: yazıda geçen madde / SUT numarası dayanak metninde ya da belgede var mı (uydurma madde olmasın);
   fatura no ve kesinti tutarı aynen geçiyor mu; yazıdaki sayılar kesinti kaydında, belgede ya da dayanakta var mı;
   "Evet" önerisi belgeye dayanıyor mu; kesin sonuç vaadi var mı. Sorun varsa yalnız o kesinti modele geri gönderilir.
4. Kod itiraz süresini (itiraz süreleri dosyası verilirse) kesinti tarihinden hesaplar; geçmiş ya da yaklaşan süreyi
   işaretler.
Çıktının "Kesintiler" sayfası "Kesinti ve İtiraz Takibi" kod bloğunun girdi sütunlarını korur.

Kullanım:
    python agent.py                                                         # örnek 3 kesinti
    python agent.py --kesintiler kesintiler.xlsx --belgeler ./belgeler --dayanak anlasma_maddeleri.txt --sureler itiraz_sureleri.csv
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import belge
import llm

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
ORNEK_TARIH = date(2026, 10, 16)
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
AI_DOLGU = PatternFill("solid", fgColor="EDE7FF")
RENK = {"Evet": "E3F5E1", "Hayır": "E8F0FE", "Belge eksik": "FFF4CE"}
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
SARI = PatternFill("solid", fgColor="FFF4CE")
YESIL = PatternFill("solid", fgColor="E3F5E1")
UST = Alignment(vertical="top", wrap_text=True)
S = {"type": "string"}
MADDE = re.compile(r"\b(?:madde|md\.?)\s*(\d+(?:\.\d+)*)|\bSUT\s*(\d+(?:\.\d+)+)|\b(Ek-\d+)\b", re.I)
VAAT = re.compile(r"\b(kesinlikle|mutlaka|şüphesiz|garanti\w*|iade edilmek zorunda|ödemek zorunda\w*|hukuki yollara|dava açıl\w*|yasal yollara)\b", re.I)


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


def sayilar(metin: str) -> set[Decimal]:
    temiz = MADDE.sub(" ", str(metin))
    temiz = re.sub(r"\b\d{1,2}:\d{2}\b", " ", temiz)
    out = set()
    for t in re.findall(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b", temiz):
        out.add(t)
    temiz = re.sub(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b", " ", temiz)
    for x in re.findall(r"\d+(?:[.,]\d+)*", temiz):
        v = para(x)
        if v is not None:
            out.add(f"{v.normalize():f}")
    return out


def maddeler(metin: str) -> set[str]:
    return {next(g for g in m.groups() if g).strip().lower() for m in MADDE.finditer(str(metin))}


def tl(x: Decimal) -> str:
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def tarih_coz(x) -> date | None:
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    try:
        return datetime.strptime(str(x or "").strip()[:10], "%d.%m.%Y").date()
    except ValueError:
        return None


# ----------------------------------------------------------------------------
# Girdi
# ----------------------------------------------------------------------------

SUTUN = {"kurum": ("sigorta sirketi", "kurum", "sigorta", "odeme kurumu"), "fatura": ("fatura no", "fatura numarasi", "takip no"),
         "tarih": ("kesinti tarihi", "tarih"), "tutar": ("kesinti tutari", "tutar"), "neden": ("kesinti nedeni", "kesinti aciklamasi", "aciklama", "neden"),
         "islem": ("islem", "islem adi", "hizmet")}


def tablo(yol: Path) -> list[list]:
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, read_only=True, data_only=True)
        s = [list(r) for r in wb.active.iter_rows(values_only=True)]
        wb.close()
    else:
        metin = yol.read_text(encoding="utf-8-sig")
        s = list(csv.reader(metin.splitlines(), delimiter=max(";\t,", key=metin.split("\n")[0].count)))
    return [r for r in s if any(c not in (None, "") for c in r)]


def kesintileri_oku(yol: Path) -> tuple[list[str], list[dict]]:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: kesinti listesi bulunamadı.")
    s = tablo(yol)
    if not s:
        raise llm.LLMHatasi(f"{yol.name}: liste boş.")
    b = [katla(c) for c in s[0]]
    idx = {a: next((i for i, x in enumerate(b) if x in es), None) for a, es in SUTUN.items()}
    eksik = [SUTUN[a][0] for a in ("kurum", "fatura", "tutar", "neden") if idx[a] is None]
    if eksik:
        raise llm.LLMHatasi(f"{yol.name}: gerekli sütun yok: {', '.join(eksik)}. Başlıklar: {s[0]}")
    out = []
    for i, r in enumerate(s[1:], 1):
        al = lambda a: None if idx[a] is None or idx[a] >= len(r) else r[idx[a]]  # noqa: E731
        t = para(al("tutar"))
        if not al("fatura") or t is None:
            continue
        out.append({"no": f"K{i}", "satir": [("" if c is None else c) for c in r], "kurum": str(al("kurum") or "").strip(),
                    "fatura": str(al("fatura")).strip(), "tarih": tarih_coz(al("tarih")), "tutar": t, "neden": str(al("neden") or "").strip(),
                    "islem": str(al("islem") or "").strip()})
    return [str(c or "") for c in s[0]], out


def sureleri_oku(yol: Path | None) -> dict[str, int]:
    if yol is None:
        return {}
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: itiraz süreleri dosyası bulunamadı.")
    return {katla(r[0]): int(para(r[1])) for r in tablo(yol)[1:] if len(r) > 1 and para(r[1]) is not None}


def takma_adlar(adlar: list[str]) -> dict[str, str]:
    return {a: f"[HASTA-{i}]" for i, a in enumerate(dict.fromkeys(x.strip() for x in adlar if x.strip()), 1)}


def maskele(metin: str, kisiler: dict[str, str]) -> str:
    for g, t in sorted(kisiler.items(), key=lambda x: -len(x[0])):
        metin = re.sub(re.escape(g), t, metin, flags=re.I)
    return llm.maskele(metin)


def belgeleri_oku(klasor: Path | None, fatura: str, kisiler: dict[str, str]) -> dict[str, str]:
    """Fatura numarasıyla başlayan dosyalar: F-201.txt, F-201_provizyon.pdf …"""
    if klasor is None or not klasor.is_dir():
        return {}
    out = {}
    for p in sorted(klasor.iterdir()):
        if p.is_file() and p.suffix.lower() in belge.DESTEKLENEN and katla(p.stem).startswith(katla(fatura)):
            out[p.name] = maskele(belge.metin_oku(p), kisiler)
    return out


# ----------------------------------------------------------------------------
# Model
# ----------------------------------------------------------------------------

SEMA = {"type": "object", "properties": {
    "oneri": {"type": "string", "enum": ["Evet", "Hayır", "Belge eksik"]}, "gerekce": S,
    "dayanak_maddeler": {"type": "array", "items": S}, "kullanilan_belgeler": {"type": "array", "items": S},
    "eksik_belgeler": {"type": "array", "items": S}, "itiraz_metni": S},
    "required": ["oneri", "gerekce", "dayanak_maddeler", "kullanilan_belgeler", "eksik_belgeler", "itiraz_metni"], "additionalProperties": False}


def denetle(k: dict, c: dict, dayanak: str) -> list[str]:
    on, sorun = k["no"], []
    kaynak = dayanak + "\n" + "\n".join(k["belgeler"].values())
    gecerli = maddeler(kaynak)
    metin = c["itiraz_metni"] + "\n" + c["gerekce"]
    yok = sorted(maddeler(metin + " " + " ".join(c["dayanak_maddeler"])) - gecerli)
    if yok:
        sorun.append(f"{on}: dayanak metninde ya da belgelerde olmayan madde / ek: {', '.join(yok)} (yalnız verilen maddelere dayanın)")
    izinli = sayilar(kaynak + " " + " ".join(str(x) for x in k["satir"])) | sayilar(f"{k['tarih']:%d.%m.%Y}" if k["tarih"] else "")
    fazla = sorted(sayilar(metin) - izinli)
    if fazla:
        sorun.append(f"{on}: kesinti kaydında, belgede ya da dayanakta olmayan sayı / tarih: {', '.join(fazla)}")
    yok_b = [b for b in c["kullanilan_belgeler"] if b not in k["belgeler"]]
    if yok_b:
        sorun.append(f"{on}: kullanılan belge {', '.join(yok_b)} verilen belgeler arasında yok ({', '.join(k['belgeler']) or 'belge yok'})")
    if c["oneri"] == "Evet" and not c["kullanilan_belgeler"]:
        sorun.append(f"{on}: 'Evet' önerisi bir belgeye dayanmalı; belge yoksa 'Belge eksik' seçin")
    if c["oneri"] == "Evet":
        if k["fatura"] not in c["itiraz_metni"]:
            sorun.append(f"{on}: itiraz metninde fatura no '{k['fatura']}' yok")
        if tl(k["tutar"]) not in c["itiraz_metni"]:
            sorun.append(f"{on}: itiraz metninde kesinti tutarı '{tl(k['tutar'])}' yok (bu biçimde yazın)")
    elif c["itiraz_metni"].strip() and c["oneri"] == "Hayır":
        sorun.append(f"{on}: itiraz önerilmiyorsa itiraz metni boş olmalı")
    if c["oneri"] == "Belge eksik" and not c["eksik_belgeler"]:
        sorun.append(f"{on}: 'Belge eksik' önerisinde eksik belgeler yazılmalı")
    if v := VAAT.search(metin):
        sorun.append(f"{on}: kesinlik / tehdit ifadesi ('{v.group(0)}'); itiraz yazısı nesnel ve belgeye dayalı olmalı")
    if re.search(r"\[(HASTA-\d+|TCKN|TELEFON|E-POSTA|IBAN)\]", c["itiraz_metni"]):
        sorun.append(f"{on}: itiraz metninde maskeli yer tutucu var; hasta bilgisi yerine fatura no kullanın")
    return sorun


def sure_durumu(k: dict, sureler: dict[str, int], bugun: date) -> tuple[date | None, str]:
    gun = sureler.get(katla(k["kurum"]))
    if gun is None or k["tarih"] is None:
        return None, "İtiraz süresi bilinmiyor" if sureler else ""
    son = k["tarih"] + timedelta(days=gun)
    kalan = (son - bugun).days
    if kalan < 0:
        return son, f"İtiraz süresi geçmiş ({son:%d.%m.%Y})"
    if kalan <= 5:
        return son, f"İtiraz süresinin bitmesine {kalan} gün var ({son:%d.%m.%Y})"
    return son, f"Son itiraz tarihi {son:%d.%m.%Y} ({kalan} gün)"


def uret(kesintiler: list[dict], dayanak: str, duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    sonuc = dict(onceki or {})
    hedef = kesintiler if duzeltme is None else [k for k in kesintiler if any(s.startswith(k["no"] + ":") for s in duzeltme)]
    for k in hedef:
        bel = "\n".join(f'<belge ad="{a}">\n{m}\n</belge>' for a, m in k["belgeler"].items()) or "(bu fatura için belge verilmedi)"
        tarih = f"{k['tarih']:%d.%m.%Y}" if k["tarih"] else "belirtilmemiş"
        mesaj = (f"<dayanak>\n{dayanak or '(dayanak metni verilmedi)'}\n</dayanak>\n"
                 f"<kesinti kurum=\"{k['kurum']}\" fatura=\"{k['fatura']}\" tarih=\"{tarih}\" tutar=\"{tl(k['tutar'])}\" islem=\"{k['islem']}\">\n"
                 f"Kesinti nedeni: {k['neden']}\n</kesinti>\n<belgeler>\n{bel}\n</belgeler>")
        d = [s.split(": ", 1)[1] for s in (duzeltme or []) if s.startswith(k["no"] + ":")]
        if d:
            mesaj += ("\n<duzeltme>\nÖnceki taslak: " + json.dumps(sonuc.get(k["no"]), ensure_ascii=False) + "\nKodun bulduğu sorunlar:\n"
                      + "\n".join(f"- {s}" for s in d) + "\n</duzeltme>")
        sonuc[k["no"]] = llm.json_iste(sistem, mesaj, SEMA)
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

def _baslik(ws, basliklar, genislik=()) -> None:
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font, h.alignment = BASLIK_DOLGU, BASLIK_YAZI, UST
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def geri_ac(s: str, kisiler: dict[str, str]) -> str:
    for g, t in kisiler.items():
        s = s.replace(t, g)
    return s


def rapor_yaz(cikti: Path, basliklar: list[str], kesintiler: list[dict], sonuc: dict, kalan: list[str], tur: int, imza: str,
              kisiler: dict[str, str]) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Kesintiler"
    ek = ["İtiraz Önerisi", "Gerekçe", "Dayanak Maddeler", "Kullanılan Belgeler", "Eksik Belgeler", "Son İtiraz Tarihi", "Süre Durumu", "Son Kontrol"]
    _baslik(ws, basliklar + ek)
    yazilar = []
    for k in kesintiler:
        c = sonuc[k["no"]]
        sor = [x.split(": ", 1)[1] for x in kalan if x.startswith(k["no"] + ":")]
        ws.append(k["satir"][:len(basliklar)] + [""] * (len(basliklar) - len(k["satir"])) +
                  [c["oneri"], geri_ac(c["gerekce"], kisiler), ", ".join(c["dayanak_maddeler"]), ", ".join(c["kullanilan_belgeler"]),
                   "\n".join(f"• {x}" for x in c["eksik_belgeler"]), k["son"], k["sure"], "; ".join(sor) or "Geçti"])
        n, b = ws.max_row, len(basliklar)
        ws.cell(n, b + 1).fill = PatternFill("solid", fgColor=RENK[c["oneri"]])
        for j in (b + 2, b + 3, b + 4, b + 5):
            ws.cell(n, j).fill = AI_DOLGU
        ws.cell(n, b + 6).number_format = "DD.MM.YYYY"
        if "geçmiş" in k["sure"] or "gün var" in k["sure"]:
            ws.cell(n, b + 7).fill = KIRMIZI if "geçmiş" in k["sure"] else SARI
        ws.cell(n, b + 8).fill = KIRMIZI if sor else YESIL
        for h in ws[n]:
            h.alignment = UST
        if c["oneri"] == "Evet":
            yazilar.append((k, geri_ac(c["itiraz_metni"], kisiler)))
    for j in range(1, len(basliklar) + len(ek) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 40 if j > len(basliklar) and j <= len(basliklar) + 5 else 14
    y = wb.create_sheet("İtiraz Yazıları")
    _baslik(y, ["Kurum", "Fatura No", "Kesinti Tutarı", "Son İtiraz Tarihi", "İtiraz Yazısı"], (16, 12, 12, 12, 110))
    for k, t in yazilar:
        y.append([k["kurum"], k["fatura"], float(k["tutar"]), k["son"], f"{t.strip()}\n\nSaygılarımızla,\n{imza}"])
        y.cell(y.max_row, 3).number_format = "#,##0.00"
        y.cell(y.max_row, 4).number_format = "DD.MM.YYYY"
        y.cell(y.max_row, 5).fill = AI_DOLGU
        for h in y[y.max_row]:
            h.alignment = UST
    kt = wb.create_sheet("Kontrol")
    kt.append([f"Düzeltme turu: {tur} · Model: {llm.kullanim_ozeti()} · {datetime.now():%d.%m.%Y %H:%M}"])
    kt.append(["Son kontrol: " + ("; ".join(kalan) if kalan else "Kod denetiminden sorunsuz geçti")])
    kt.cell(2, 1).fill = KIRMIZI if kalan else YESIL
    kt.append(["Önemli: İtiraz yazıları taslaktır; göndermeden önce faturalama ve tıbbi birimle kontrol edin. Yazılar yalnız verilen dayanak "
               "metnine ve belgelere dayanır; madde numaraları kod tarafından dayanak metniyle karşılaştırılmıştır. İtiraz tarihini gönderdiğinizde "
               "'İtiraz Tarihi' sütununa yazın; dosya 'Kesinti ve İtiraz Takibi' kod bloğunda kullanılabilir."])
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)
    txt = cikti.with_suffix(".txt")
    txt.write_text("\n\n".join(f"--- {k['kurum']} · {k['fatura']} · {tl(k['tutar'])} TL ---\n{t.strip()}\n\nSaygılarımızla,\n{imza}" for k, t in yazilar),
                   encoding="utf-8")
    return txt


def calistir(kesinti_yolu: Path, cikti: Path, belge_klasoru: Path | None = None, dayanak_yolu: Path | None = None, sure_yolu: Path | None = None,
             bugun: date = ORNEK_TARIH, gizle: list[str] | None = None, imza: str = "Anlaşmalı Kurumlar Birimi", tur: int = 2,
             evet: bool = False) -> dict:
    kisiler = takma_adlar(gizle or [])
    basliklar, kesintiler = kesintileri_oku(kesinti_yolu)
    sureler = sureleri_oku(sure_yolu)
    if dayanak_yolu is not None and not dayanak_yolu.exists():
        raise llm.LLMHatasi(f"{dayanak_yolu}: dayanak dosyası bulunamadı.")
    dayanak = maskele(belge.metin_oku(dayanak_yolu), kisiler) if dayanak_yolu else ""
    for k in kesintiler:
        k["belgeler"] = belgeleri_oku(belge_klasoru, k["fatura"], kisiler)
        k["son"], k["sure"] = sure_durumu(k, sureler, bugun)
    print(f"[OK] {len(kesintiler)} kesinti · belgeli: {sum(bool(k['belgeler']) for k in kesintiler)} · dayanak: {len(dayanak):,} karakter")
    llm.onay_al(f"{len(kesintiler)} kesinti kaydı, destek belgeleri ve dayanak metni gönderilecek. Hasta adları (--gizle), TCKN, telefon, e-posta "
                "ve IBAN maskeli; sağlık bilgisi içeren belgelerde yerel model (Ollama) kullanmayı düşünün.", evet)
    sonuc, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(kesintiler, dayanak, d, o),
                                             lambda x: [s for k in kesintiler for s in denetle(k, x[k["no"]], dayanak)], tur)
    txt = rapor_yaz(cikti, basliklar, kesintiler, sonuc, kalan, yapilan, imza, kisiler)
    return {"kesintiler": kesintiler, "sonuc": sonuc, "kalan": kalan, "tur": yapilan, "txt": txt}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    p = argparse.ArgumentParser(description="Fatura kesintileri için belgeye ve anlaşma / mevzuat metnine dayanan itiraz yazısı taslakları hazırlar.")
    p.add_argument("--kesintiler", type=Path, default=ORNEK / "kesintiler.csv", help="Kesinti listesi (.xlsx/.csv): Sigorta Şirketi, Fatura No, Kesinti Tarihi, Kesinti Tutarı, Kesinti Nedeni")
    p.add_argument("--belgeler", type=Path, default=None, help="Destek belgeleri klasörü; dosya adı fatura no ile başlar (F-201.txt, F-201_epikriz.pdf)")
    p.add_argument("--dayanak", type=Path, default=None, help="Anlaşma / SUT maddeleri metni (.txt/.docx/.pdf)")
    p.add_argument("--sureler", type=Path, default=None, help="İtiraz süreleri (.csv: Sigorta Şirketi; İtiraz Süresi (gün))")
    p.add_argument("--tarih", default=None, help="Bugünün tarihi GG.AA.YYYY (varsayılan bugün; örnekte 16.10.2026)")
    p.add_argument("--gizle", nargs="*", default=[], metavar="AD", help="Sağlayıcıya gönderilmeyecek hasta / kişi adları")
    p.add_argument("--imza", default="Anlaşmalı Kurumlar Birimi")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "kesinti_itirazlari.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    ornek = a.kesintiler == ORNEK / "kesintiler.csv"
    bugun = tarih_coz(a.tarih) if a.tarih else (ORNEK_TARIH if ornek else date.today())
    if bugun is None:
        print(f"[X] --tarih '{a.tarih}' okunamadı (GG.AA.YYYY)")
        return 2
    try:
        s = calistir(a.kesintiler, a.cikti, a.belgeler or (ORNEK / "belgeler" if ornek else None), a.dayanak or (ORNEK / "dayanaklar.txt" if ornek else None),
                     a.sureler or (ORNEK / "itiraz_sureleri.csv" if ornek else None), bugun, a.gizle, a.imza, a.tur, a.evet)
    except (llm.LLMHatasi, belge.BelgeHatasi) as h:
        print(f"[X] {h}")
        return 1
    for k in s["kesintiler"]:
        c = s["sonuc"][k["no"]]
        print(f"[OK] {k['fatura']} · {k['kurum']} · {tl(k['tutar'])} TL · öneri: {c['oneri']} · {k['sure']}")
    for x in s["kalan"]:
        print(f"[!] Kalan: {x}")
    print(f"[OK] Rapor: {a.cikti.resolve()} (+ {s['txt'].name})")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
