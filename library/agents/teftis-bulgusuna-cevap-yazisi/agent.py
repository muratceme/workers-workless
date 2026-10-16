"""
Teftiş Bulgusuna Cevap Yazısı — Workers / Workless AI Agent
Şube Bankacılığı › Şube Operasyon Yöneticisi

Teftiş / iç kontrol bulgularına, şubenin kendi notlarına dayanan cevap yazısı taslağı hazırlar ve bulgu takip tablosunu
günceller (kapanış tarihi, termin, aksiyon sahibi).
1. Kod bulgu listesini ("Kontrol Bulgu Takibi" kod bloğunun girdi biçimi) ve şube notlarını ("Bulgu: No" başlıklı bloklar)
   okur; telefon, e-posta, kimlik numarası ve --gizle adları maskelenir.
2. Model her bulgu için durum (Tamamlandı / Devam ediyor / Planlandı / İtiraz), kök neden, alınan ve planlanan aksiyonları,
   kapanış / termin tarihini, aksiyon sahibini ve cevap paragrafını yazar.
3. Kod denetler: her bulgunun notu ve cevabı var mı; tarih ve sayılar notta geçiyor mu; "Tamamlandı" için notta tamamlama
   ifadesi ve kapanış tarihi var mı; "Planlandı / Devam ediyor" için notta termin var mı; termin açılıştan sonra mı;
   notta olmayan aksiyon "yapıldı" diye yazılmış mı. Sorun varsa yalnız o bulgu modele geri gönderilir.
4. Kod yazıyı sabit düzenle kurar ve takip tablosunu günceller.

Kullanım:
    python agent.py                                                         # örnek 3 bulgu
    python agent.py --bulgular bulgular.xlsx --notlar sube_notlari.txt --sube "Kadıköy Şubesi" --imza "Şube Operasyon Yöneticisi"
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from datetime import date, datetime
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
RENK = {"Tamamlandı": "E3F5E1", "Devam ediyor": "FFF4CE", "Planlandı": "E8F0FE", "İtiraz": "FDE2E1"}
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
YESIL = PatternFill("solid", fgColor="E3F5E1")
UST = Alignment(vertical="top", wrap_text=True)
S = {"type": "string"}
TAMAM = re.compile(r"\b(tamamland|yapild|giderild|duzeltild|kapatild|duzenlen|imzalatild|verildi|alindi)\w*")
GELECEK = re.compile(r"\b(planliyor|planlaniyor|planlandi|yapilacak|tamamlanacak|edilecek|alinacak|verilecek|devreye alinacak|bekleniyor|surecek)\w*")


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9%]+", " ", str(s or "").replace("İ", "i").replace("I", "ı").translate(_TR).lower()).split())


def para(x) -> Decimal | None:
    s = re.sub(r"[^\d,.\-]", "", str(x or ""))
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return Decimal(s) if s not in ("", "-", ".") else None
    except InvalidOperation:
        return None


def sayilar(metin: str) -> set[Decimal]:
    temiz = re.sub(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b|\b[A-Z]{1,4}-[\w-]+\b", " ", str(metin))
    return {v for x in re.findall(r"\d+(?:[.,]\d+)*", temiz) if (v := para(x)) is not None}


def tarihler(metin: str) -> set[str]:
    return set(re.findall(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b", str(metin)))


def tarih_coz(x) -> date | None:
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    m = re.search(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4})\b", str(x or ""))
    try:
        return date(int(m[3]), int(m[2]), int(m[1])) if m else None
    except ValueError:
        return None


# ----------------------------------------------------------------------------
# Girdi
# ----------------------------------------------------------------------------

def bulgulari_oku(yol: Path) -> tuple[list[str], list[dict]]:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: bulgu listesi bulunamadı.")
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, read_only=True, data_only=True)
        s = [list(r) for r in wb.active.iter_rows(values_only=True)]
        wb.close()
    else:
        metin = yol.read_text(encoding="utf-8-sig")
        s = list(csv.reader(metin.splitlines(), delimiter=max(";\t,", key=metin.split("\n")[0].count)))
    s = [r for r in s if any(c not in (None, "") for c in r)]
    bas = [str(c or "") for c in s[0]]
    b = [katla(c) for c in bas]
    if "bulgu no" not in b or "bulgu" not in b:
        raise llm.LLMHatasi(f"{yol.name}: 'Bulgu No' ve 'Bulgu' sütunları gerekli.")
    out = []
    for r in s[1:]:
        r = list(r) + [None] * (len(bas) - len(r))
        d = dict(zip(b, r))
        if d.get("bulgu no"):
            out.append({"no": str(d["bulgu no"]).strip(), "bulgu": str(d.get("bulgu") or ""), "onem": str(d.get("onem") or ""),
                        "acilis": tarih_coz(d.get("acilis tarihi")), "satir": r})
    return bas, out


def notlari_oku(yol: Path, kisiler: dict[str, str]) -> dict[str, str]:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: not dosyası bulunamadı.")
    out, son = {}, None
    for s in belge.metin_oku(yol).splitlines():
        m = re.match(r"^\s*Bulgu\s*:\s*(\S+)\s*$", s, re.I)
        if m:
            son = m[1]
            out[son] = ""
        elif son and s.strip():
            out[son] = (out[son] + " " + maskele(s.strip(), kisiler)).strip()
    return out


def takma_adlar(adlar: list[str]) -> dict[str, str]:
    return {a: f"[KİŞİ-{i}]" for i, a in enumerate(dict.fromkeys(x.strip() for x in adlar if x.strip()), 1)}


def maskele(metin: str, kisiler: dict[str, str]) -> str:
    for g, t in sorted(kisiler.items(), key=lambda x: -len(x[0])):
        metin = re.sub(re.escape(g), t, metin, flags=re.I)
    return llm.maskele(metin)


# ----------------------------------------------------------------------------
# Model
# ----------------------------------------------------------------------------

SEMA = {"type": "object", "properties": {
    "durum": {"type": "string", "enum": ["Tamamlandı", "Devam ediyor", "Planlandı", "İtiraz"]}, "kok_neden": S,
    "alinan_aksiyonlar": {"type": "array", "items": S}, "planlanan_aksiyonlar": {"type": "array", "items": S},
    "kapanis_tarihi": S, "termin": S, "aksiyon_sahibi": S, "cevap": S},
    "required": ["durum", "kok_neden", "alinan_aksiyonlar", "planlanan_aksiyonlar", "kapanis_tarihi", "termin", "aksiyon_sahibi", "cevap"],
    "additionalProperties": False}


def denetle(b: dict, c: dict) -> list[str]:
    on, notu, sorun = b["no"], b["not"], []
    metin = " ".join([c["kok_neden"], c["cevap"], *c["alinan_aksiyonlar"], *c["planlanan_aksiyonlar"]])
    fazla_t = sorted(tarihler(metin + " " + c["kapanis_tarihi"] + " " + c["termin"]) - tarihler(notu))
    if fazla_t:
        sorun.append(f"{on}: notta olmayan tarih: {', '.join(fazla_t)}")
    fazla = sorted(sayilar(metin) - sayilar(notu) - sayilar(b["bulgu"]), key=float)
    if fazla:
        sorun.append(f"{on}: notta ya da bulguda olmayan sayı: {', '.join(f'{x:g}' for x in fazla)}")
    eksik_t = sorted(tarihler(notu) - tarihler(metin + " " + c["kapanis_tarihi"] + " " + c["termin"]))
    if eksik_t:
        sorun.append(f"{on}: nottaki tarih cevapta yok: {', '.join(eksik_t)} (bu tarihli işi atlamayın)")
    if (g := GELECEK.search(katla(notu))) and (c["durum"] == "Tamamlandı" or not c["planlanan_aksiyonlar"]):
        sorun.append(f"{on}: notta henüz yapılmamış iş var ('{g.group(0)}'); durum 'Tamamlandı' olamaz ve planlanan aksiyon yazılmalı")
    if c["durum"] == "Tamamlandı":
        if not TAMAM.search(katla(notu)):
            sorun.append(f"{on}: notta tamamlandığını gösteren ifade yok; 'Tamamlandı' yerine 'Devam ediyor' / 'Planlandı' seçin")
        if not c["kapanis_tarihi"].strip():
            sorun.append(f"{on}: 'Tamamlandı' için notta yazan kapanış tarihi gerekli")
        if c["planlanan_aksiyonlar"]:
            sorun.append(f"{on}: planlanan aksiyon varken durum 'Tamamlandı' olamaz")
    if c["durum"] in ("Devam ediyor", "Planlandı") and not c["termin"].strip():
        sorun.append(f"{on}: '{c['durum']}' için notta yazan termin gerekli; notta yoksa durumu belirtip termini boş bırakmayın, şubeye sorun")
    t = tarih_coz(c["termin"])
    if t and b["acilis"] and t < b["acilis"]:
        sorun.append(f"{on}: termin ({t:%d.%m.%Y}) bulgu açılışından önce")
    if c["durum"] == "İtiraz" and len(c["kok_neden"].strip()) < 10:
        sorun.append(f"{on}: itiraz için nottaki gerekçeyi yazın")
    if not c["cevap"].strip():
        sorun.append(f"{on}: cevap paragrafı boş")
    return sorun


def uret(bulgular: list[dict], duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    sonuc = dict(onceki or {})
    hedef = bulgular if duzeltme is None else [b for b in bulgular if any(s.startswith(b["no"] + ":") for s in duzeltme)]
    for b in hedef:
        ac = f"{b['acilis']:%d.%m.%Y}" if b["acilis"] else "-"
        mesaj = f"<bulgu no=\"{b['no']}\" onem=\"{b['onem']}\" acilis=\"{ac}\">{b['bulgu']}</bulgu>\n<sube_notu>{b['not']}</sube_notu>"
        d = [s.split(": ", 1)[1] for s in (duzeltme or []) if s.startswith(b["no"] + ":")]
        if d:
            mesaj += ("\n<duzeltme>\nÖnceki cevap: " + json.dumps(sonuc.get(b["no"]), ensure_ascii=False) + "\nKodun bulduğu sorunlar:\n"
                      + "\n".join(f"- {s}" for s in d) + "\n</duzeltme>")
        sonuc[b["no"]] = llm.json_iste(sistem, mesaj, SEMA)
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


def geri_ac(s, kisiler: dict[str, str]):
    if isinstance(s, str):
        for g, t in kisiler.items():
            s = s.replace(t, g)
        return s
    if isinstance(s, list):
        return [geri_ac(x, kisiler) for x in s]
    if isinstance(s, dict):
        return {k: geri_ac(v, kisiler) for k, v in s.items()}
    return s


# ----------------------------------------------------------------------------
# Çıktı
# ----------------------------------------------------------------------------

def yazi(bulgular: list[dict], sonuc: dict, sube: str, imza: str, eksik: list[dict]) -> str:
    l = [f"Konu: {sube or 'Şubemiz'} teftiş / kontrol bulgularına ilişkin cevaplarımız", "", "Teftiş Kurulu Başkanlığına / İç Kontrol Birimine,", "",
         f"{sube or 'Şubemizde'} tespit edilen bulgulara ilişkin açıklamalarımız ve aldığımız aksiyonlar aşağıdadır.", ""]
    for b in bulgular:
        c = sonuc[b["no"]]
        l.append(f"{b['no']} — {b['bulgu']}")
        l.append(f"Durum: {c['durum']}" + (f" (kapanış {c['kapanis_tarihi']})" if c["kapanis_tarihi"] else "") + (f" · Termin: {c['termin']}" if c["termin"] else "")
                 + (f" · Sorumlu: {c['aksiyon_sahibi']}" if c["aksiyon_sahibi"] else ""))
        l.append(c["cevap"].strip())
        l.append("")
    for b in eksik:
        l += [f"{b['no']} — {b['bulgu']}", "Cevap hazırlanmadı: şube notu yok.", ""]
    l += ["Bilgilerinize arz ederiz.", "", "Saygılarımızla,", imza]
    return "\n".join(l)


def rapor_yaz(cikti: Path, bas: list[str], bulgular: list[dict], eksik: list[dict], sonuc: dict, kalan: list[str], tur: int, sube: str, imza: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Bulgular"
    ek = ["Durum", "Kök Neden", "Alınan Aksiyonlar", "Planlanan Aksiyonlar", "Cevap", "Son Kontrol"]
    ws.append(bas + ek)
    for h in ws[1]:
        h.fill, h.font, h.alignment = BASLIK_DOLGU, BASLIK_YAZI, UST
    kb = [katla(x) for x in bas]
    for b in bulgular + eksik:
        r = list(b["satir"])
        c = sonuc.get(b["no"])
        if c:
            for alan, deger in (("termin", tarih_coz(c["termin"])), ("kapanis tarihi", tarih_coz(c["kapanis_tarihi"])), ("aksiyon sahibi", c["aksiyon_sahibi"])):
                if alan in kb and deger:
                    r[kb.index(alan)] = deger
        sor = [x.split(": ", 1)[1] for x in kalan if x.startswith(b["no"] + ":")]
        ws.append(r + ([c["durum"], c["kok_neden"], "\n".join(f"• {x}" for x in c["alinan_aksiyonlar"]), "\n".join(f"• {x}" for x in c["planlanan_aksiyonlar"]),
                        c["cevap"], "; ".join(sor) or "Geçti"] if c else ["", "", "", "", "Şube notu yok", ""]))
        n = ws.max_row
        for j, v in enumerate(r, 1):
            if isinstance(v, (date, datetime)):
                ws.cell(n, j).number_format = "DD.MM.YYYY"
        if c:
            ws.cell(n, len(bas) + 1).fill = PatternFill("solid", fgColor=RENK[c["durum"]])
            for j in range(len(bas) + 2, len(bas) + 6):
                ws.cell(n, j).fill = AI_DOLGU
            ws.cell(n, len(bas) + 6).fill = KIRMIZI if sor else YESIL
        for h in ws[n]:
            h.alignment = UST
    for j in range(1, len(bas) + len(ek) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 40 if j in (2, len(bas) + 2, len(bas) + 3, len(bas) + 4, len(bas) + 5) else 13
    y = wb.create_sheet("Cevap Yazısı")
    y.column_dimensions["A"].width = 120
    metin = yazi(bulgular, sonuc, sube, imza, eksik)
    for s in metin.splitlines():
        y.append([s])
        y.cell(y.max_row, 1).alignment = UST
    k = wb.create_sheet("Kontrol")
    k.append([f"Düzeltme turu: {tur} · Model: {llm.kullanim_ozeti()} · {datetime.now():%d.%m.%Y %H:%M}"])
    k.append(["Son kontrol: " + ("; ".join(kalan) if kalan else "Kod denetiminden sorunsuz geçti")])
    k.cell(2, 1).fill = KIRMIZI if kalan else YESIL
    k.append(["Önemli: Cevaplar yalnız şube notlarına dayanır. Göndermeden önce şube müdürü onaylamalı; 'Tamamlandı' denen her aksiyonun kanıtı "
              "(tutanak, eğitim katılım listesi, sistem kaydı) dosyada olmalıdır. 'Bulgular' sayfası 'Kontrol Bulgu Takibi' kod bloğuna verilebilir."])
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)
    txt = cikti.with_suffix(".txt")
    txt.write_text(metin, encoding="utf-8")
    return txt


def calistir(bulgu_yolu: Path, not_yolu: Path, cikti: Path, sube: str = "", imza: str = "Şube Operasyon Yöneticisi", gizle: list[str] | None = None,
             tur: int = 2, evet: bool = False) -> dict:
    kisiler = takma_adlar(gizle or [])
    bas, bulgular = bulgulari_oku(bulgu_yolu)
    notlar = notlari_oku(not_yolu, kisiler)
    eksik = [b for b in bulgular if not notlar.get(b["no"])]
    bulgular = [b | {"not": notlar[b["no"]]} for b in bulgular if notlar.get(b["no"])]
    print(f"[OK] {len(bulgular) + len(eksik)} bulgu · notu olan {len(bulgular)} · notu olmayan {len(eksik)}")
    sonuc, kalan, yapilan = {}, [], 0
    if bulgular:
        llm.onay_al(f"{len(bulgular)} bulgu metni ve şube notları gönderilecek (telefon, e-posta, TCKN maskeli"
                    f"{f'; {len(kisiler)} ad takma adla' if kisiler else ''}).", evet)
        sonuc, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(bulgular, d, o), lambda x: [s for b in bulgular for s in denetle(b, x[b["no"]])], tur)
        sonuc = geri_ac(sonuc, kisiler)
    txt = rapor_yaz(cikti, bas, bulgular, eksik, sonuc, kalan, yapilan, sube, imza)
    return {"bulgular": bulgular, "eksik": eksik, "sonuc": sonuc, "kalan": kalan, "tur": yapilan, "txt": txt}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    p = argparse.ArgumentParser(description="Teftiş / iç kontrol bulgularına şube notlarına dayanan cevap yazısı hazırlar ve takip tablosunu günceller.")
    p.add_argument("--bulgular", type=Path, default=ORNEK / "bulgular.csv", help="Bulgu listesi (.xlsx/.csv): Bulgu No, Bulgu, Önem, Açılış Tarihi …")
    p.add_argument("--notlar", type=Path, default=ORNEK / "sube_notlari.txt", help="Şube notları (.txt/.docx): her blok 'Bulgu: No' satırıyla başlar")
    p.add_argument("--sube", default="", help="Şube adı")
    p.add_argument("--imza", default="Şube Operasyon Yöneticisi")
    p.add_argument("--gizle", nargs="*", default=[], metavar="AD", help="Notlarda geçen ve gönderilmeyecek personel / müşteri adları")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "bulgu_cevaplari.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    ornek = a.bulgular == ORNEK / "bulgular.csv"
    try:
        s = calistir(a.bulgular, a.notlar, a.cikti, a.sube or ("Kadıköy Şubesi" if ornek else ""), a.imza, a.gizle, a.tur, a.evet)
    except (llm.LLMHatasi, belge.BelgeHatasi) as h:
        print(f"[X] {h}")
        return 1
    for b in s["bulgular"]:
        c = s["sonuc"][b["no"]]
        print(f"[OK] {b['no']} · {c['durum']}" + (f" · termin {c['termin']}" if c["termin"] else "") + (f" · kapanış {c['kapanis_tarihi']}" if c["kapanis_tarihi"] else ""))
    for b in s["eksik"]:
        print(f"[!] {b['no']}: şube notu yok, cevap hazırlanmadı")
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    print(f"[OK] Rapor: {a.cikti.resolve()} (+ {s['txt'].name})")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
