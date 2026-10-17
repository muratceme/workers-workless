"""
Denetim Raporu Bulgu Aktarımı — Gerisi Bende AI Agent
Tekstil Uygunluk › Sosyal Uygunluk Uzmanı (Sürdürülebilirlik Uzmanı da kullanır)

Sosyal uygunluk / çevre / müşteri denetim raporundaki bulguları düzeltici faaliyet takip tablosuna aktarır.
1. Kod rapor metnini okur; telefon, e-posta ve kimlik numaraları maskelenir; --gizle ile verilen adlar takma adla gönderilir.
2. Model her bulgu için konu, bulgu metni, rapordaki derece, rapordan alıntı (kanıt), düzeltici faaliyet önerisi ve sorumlu
   görev önerisi yazar; denetim türü, tarihi ve tesisi raporun başından alır.
3. Kod çıkarımı denetler: alıntı raporda aynen geçiyor mu; derece raporda o bulgu için yazan derece mi; sayılar raporda var
   mı; denetim tarihi raporda yazıyor mu. Sorun varsa modele geri gönderilir (en fazla --tur kez).
4. Kod termini dereceye göre hesaplar (denetim tarihi + gün; --termin ile değişir), sıfır tolerans konularını (çocuk / genç
   işçi, zorla çalıştırma, ücret ödenmemesi, rüşvet, kilitli çıkış…) raporun derecesinden bağımsız işaretler; gözlemleri
   takip listesine almaz.
Çıktının ilk sayfası "Denetim Bulgu ve Düzeltici Faaliyet Takibi" kod bloğunun girdi sütunlarındadır.

Kullanım:
    python agent.py                                                         # örnek rapor (5 bulgu)
    python agent.py --rapor denetim.pdf --termin "kritik=7,majör=30,minör=60" --onek BSCI-26
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook
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
RENK = {"Kritik": "FDE2E1", "Majör": "FFF4CE", "Minör": "E8F0FE", "Gözlem": "E3F5E1", "Belirtilmemiş": "FFF4CE"}
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
YESIL = PatternFill("solid", fgColor="E3F5E1")
UST = Alignment(vertical="top", wrap_text=True)
S = {"type": "string"}
DERECE = ["Kritik", "Majör", "Minör", "Gözlem", "Belirtilmemiş"]
ESDEGER = {"kritik": "Kritik", "critical": "Kritik", "zero tolerance": "Kritik", "sifir tolerans": "Kritik", "major": "Majör", "majör": "Majör",
           "minor": "Minör", "minör": "Minör", "gozlem": "Gözlem", "observation": "Gözlem", "iyi uygulama": "Gözlem", "good practice": "Gözlem"}
VARSAYILAN_TERMIN = {"Kritik": 7, "Majör": 30, "Minör": 60, "Belirtilmemiş": 30}
SIFIR = re.compile(r"\b(cocuk isci|cocuk calis|genc isci|18 yas\w* alti|1[4-7] yas\w*|zorla calis|zorunlu calis|kimlik\w*(?: \w+){0,4} (alikon|tutul|el kon)|"
                   r"borc esareti|ucret\w* (odenmedi|odenmemis|odenmiyor)|asgari ucret\w* alti|rusvet|sahte kayit|cift kayit|kilitli (cikis|kapi)|"
                   r"fiziksel (siddet|ceza)|taciz)")


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9%]+", " ", str(s or "").replace("İ", "i").replace("I", "ı").translate(_TR).lower()).split())


def metinde_var(alinti: str, metin: str) -> bool:
    a = katla(alinti)
    return bool(a) and a in katla(metin)


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
    temiz = re.sub(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b|\b\d{1,2}:\d{2}\b", " ", str(metin))
    temiz = re.sub(r"(?m)^\s*\d{1,2}[.)]\s", " ", temiz)
    return {v for x in re.findall(r"\d+(?:[.,]\d+)*", temiz) if (v := para(x)) is not None}


def tarih_coz(x) -> date | None:
    m = re.search(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4})\b", str(x or ""))
    if not m:
        return None
    try:
        return date(int(m[3]), int(m[2]), int(m[1]))
    except ValueError:
        return None


def termin_oku(metin: str) -> dict[str, int]:
    out = dict(VARSAYILAN_TERMIN)
    for parca in filter(None, (x.strip() for x in metin.split(","))):
        ad, _, gun = parca.partition("=")
        anahtar = ESDEGER.get(katla(ad), {"belirtilmemis": "Belirtilmemiş"}.get(katla(ad)))
        if anahtar not in out or not gun.strip().isdigit():
            raise ValueError(f"'{parca}' okunamadı (ör. kritik=7,majör=30,minör=60)")
        out[anahtar] = int(gun)
    return out


def takma_adlar(adlar: list[str]) -> dict[str, str]:
    return {a: f"[KİŞİ-{i}]" for i, a in enumerate(dict.fromkeys(x.strip() for x in adlar if x.strip()), 1)}


def maskele(metin: str, kisiler: dict[str, str]) -> str:
    for g, t in sorted(kisiler.items(), key=lambda x: -len(x[0])):
        metin = re.sub(re.escape(g), t, metin, flags=re.I)
    return llm.maskele(metin)


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
# Model
# ----------------------------------------------------------------------------

BULGU = {"type": "object", "properties": {"konu": S, "bulgu": S, "derece": {"type": "string", "enum": DERECE}, "rapordaki_derece": S, "alinti": S,
                                          "duzeltici_faaliyet": S, "sorumlu": S},
         "required": ["konu", "bulgu", "derece", "rapordaki_derece", "alinti", "duzeltici_faaliyet", "sorumlu"], "additionalProperties": False}
SEMA = {"type": "object", "properties": {"denetim": S, "denetim_tarihi": S, "tesis": S, "bulgular": {"type": "array", "items": BULGU},
                                         "notlar": {"type": "array", "items": S}},
        "required": ["denetim", "denetim_tarihi", "tesis", "bulgular", "notlar"], "additionalProperties": False}


def denetle(c: dict, metin: str) -> list[str]:
    sorun, say = [], sayilar(metin)
    t = tarih_coz(c["denetim_tarihi"])
    if t is None or f"{t:%d.%m.%Y}" not in metin:
        sorun.append(f"Rapor: denetim tarihi '{c['denetim_tarihi']}' raporda GG.AA.YYYY olarak yazmıyor")
    if c["tesis"].strip() and not metinde_var(c["tesis"], metin):
        sorun.append(f"Rapor: tesis '{c['tesis'][:40]}' raporda aynen geçmiyor")
    if not c["bulgular"] and not c["notlar"]:
        sorun.append("Rapor: bulgu yok; raporda bulgu yoksa notlara yazın")
    for i, b in enumerate(c["bulgular"], 1):
        on = f"Bulgu {i}"
        if not metinde_var(b["alinti"], metin):
            sorun.append(f"{on}: alıntı '{b['alinti'][:50]}' raporda aynen geçmiyor")
        if b["rapordaki_derece"].strip():
            if not metinde_var(b["rapordaki_derece"], metin):
                sorun.append(f"{on}: rapordaki derece '{b['rapordaki_derece']}' raporda geçmiyor")
            elif ESDEGER.get(katla(b["rapordaki_derece"]), "Belirtilmemiş") != b["derece"]:
                sorun.append(f"{on}: derece '{b['derece']}' rapordaki dereceyle ('{b['rapordaki_derece']}') aynı değil; raporun derecesini değiştirmeyin")
        elif b["derece"] != "Belirtilmemiş":
            sorun.append(f"{on}: raporda derece yazmıyorsa 'Belirtilmemiş' seçin")
        fazla = sorted(sayilar(b["bulgu"]) - say, key=float)
        if fazla:
            sorun.append(f"{on}: bulgu metninde raporda olmayan sayı: {', '.join(f'{x:g}' for x in fazla)}")
        if b["derece"] != "Gözlem" and not b["duzeltici_faaliyet"].strip():
            sorun.append(f"{on}: düzeltici faaliyet önerisi boş")
    return sorun


def degerlendir(c: dict, terminler: dict[str, int]) -> list[dict]:
    """Kod: termin, sıfır tolerans işareti, takip listesine girme."""
    t = tarih_coz(c["denetim_tarihi"])
    out = []
    for b in c["bulgular"]:
        sifir = SIFIR.search(katla(b["bulgu"] + " " + b["alinti"] + " " + b["konu"]))
        uyari = []
        if sifir:
            uyari.append(f"Sıfır tolerans konusu ({b['konu']}): raporda {b['derece']} yazsa da derhâl aksiyon ve müşteri bildirimi "
                         "gerekebilir; müşterinizin politikasını kontrol edin")
        if b["derece"] == "Belirtilmemiş":
            uyari.append("Raporda derece yazmıyor; denetim kuruluşuna teyit ettirin")
        takip = b["derece"] != "Gözlem"
        gun = terminler["Kritik"] if sifir else terminler.get(b["derece"])
        out.append(b | {"termin": t + timedelta(days=gun) if takip and t and gun is not None else None, "sifir": bool(sifir), "uyari": uyari, "takip": takip})
    return out


def uret(metin: str, duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    mesaj = f"<rapor>\n{metin}\n</rapor>"
    if duzeltme:
        mesaj += ("\n<duzeltme>\nÖnceki çıkarım: " + json.dumps(onceki, ensure_ascii=False) + "\nKodun bulduğu sorunlar:\n"
                  + "\n".join(f"- {s}" for s in duzeltme) + "\n</duzeltme>")
    return llm.json_iste(sistem, mesaj, SEMA)


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

TAKIP = ["Bulgu No", "Denetim", "Denetim Tarihi", "Tesis", "Konu", "Bulgu", "Derece", "Düzeltici Faaliyet", "Sorumlu", "Termin", "Durum",
         "Tamamlanma Tarihi", "Kanıt"]


def _baslik(ws, basliklar, genislik) -> None:
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font, h.alignment = BASLIK_DOLGU, BASLIK_YAZI, UST
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def rapor_yaz(cikti: Path, c: dict, bulgular: list[dict], kalan: list[str], tur: int, onek: str, kaynak: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Bulgular"
    _baslik(ws, TAKIP, (11, 22, 11, 26, 22, 60, 9, 50, 20, 11, 10, 11, 10))
    t = tarih_coz(c["denetim_tarihi"])
    d = wb.create_sheet("Ayrıntı")
    _baslik(d, ["Bulgu No", "Derece", "Rapordaki Derece", "Rapordan Alıntı", "Sıfır Tolerans", "Uyarılar"], (11, 9, 12, 80, 9, 70))
    n = 0
    for b in bulgular:
        no = f"{onek}-{n + 1:02d}" if b["takip"] else "Gözlem"
        if b["takip"]:
            n += 1
            ws.append([no, c["denetim"], t, c["tesis"], b["konu"], b["bulgu"], b["derece"] if b["derece"] != "Belirtilmemiş" else "", b["duzeltici_faaliyet"],
                       b["sorumlu"], b["termin"], "Açık", None, None])
            r = ws.max_row
            ws.cell(r, 3).number_format = ws.cell(r, 10).number_format = "DD.MM.YYYY"
            for j in (5, 6, 8, 9):
                ws.cell(r, j).fill = AI_DOLGU
            ws.cell(r, 7).fill = PatternFill("solid", fgColor=RENK[b["derece"]])
            if b["sifir"]:
                ws.cell(r, 10).fill = KIRMIZI
            for h in ws[r]:
                h.alignment = UST
        d.append([no, b["derece"], b["rapordaki_derece"], b["alinti"], "Evet" if b["sifir"] else "", "\n".join(b["uyari"])])
        if b["sifir"]:
            d.cell(d.max_row, 5).fill = KIRMIZI
        for h in d[d.max_row]:
            h.alignment = UST
    k = wb.create_sheet("Kontrol")
    k.append([f"Kaynak: {kaynak} · Düzeltme turu: {tur} · Model: {llm.kullanim_ozeti()} · {datetime.now():%d.%m.%Y %H:%M}"])
    k.append(["Son kontrol: " + ("; ".join(kalan) if kalan else "Kod denetiminden sorunsuz geçti")])
    k.cell(2, 1).fill = KIRMIZI if kalan else YESIL
    for x in c["notlar"]:
        k.append([f"Model notu: {x}"])
    k.append(["Önemli: Konu, bulgu özeti, düzeltici faaliyet ve sorumlu yapay zekâ önerisidir (mor); derece raporun kendisinden alınır ve kod "
              "değiştirilmediğini denetler. Terminler dereceye göre kodla hesaplanmıştır; müşterinizin / denetim kuruluşunun belirlediği süreler "
              "farklıysa --termin ile verin. 'Bulgular' sayfası 'Denetim Bulgu ve Düzeltici Faaliyet Takibi' kod bloğuna verilebilir."])
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(rapor: Path, cikti: Path, terminler: dict[str, int] | None = None, onek: str = "B", gizle: list[str] | None = None, tur: int = 2,
             evet: bool = False) -> dict:
    if not rapor.exists():
        raise belge.BelgeHatasi(f"{rapor}: rapor bulunamadı.")
    kisiler = takma_adlar(gizle or [])
    metin = maskele(belge.metin_oku(rapor), kisiler)
    print(f"[OK] Rapor {len(metin):,} karakter")
    llm.onay_al(f"'{rapor.name}' denetim raporunun metni gönderilecek (telefon, e-posta, TCKN maskeli"
                f"{f'; {len(kisiler)} ad takma adla' if kisiler else ''}).", evet)
    c, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(metin, d, o), lambda x: denetle(x, metin), tur)
    c = geri_ac(c, kisiler)
    bulgular = degerlendir(c, terminler or VARSAYILAN_TERMIN)
    rapor_yaz(cikti, c, bulgular, kalan, yapilan, onek, rapor.name)
    return {"c": c, "bulgular": bulgular, "kalan": kalan, "tur": yapilan}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    p = argparse.ArgumentParser(description="Denetim raporu bulgularını düzeltici faaliyet takip tablosuna aktarır.")
    p.add_argument("--rapor", type=Path, default=ORNEK / "sosyal_denetim_raporu.txt", help="Denetim raporu (.pdf, .docx, .txt)")
    p.add_argument("--termin", default="", help="Dereceye göre termin, gün (varsayılan kritik=7,majör=30,minör=60,belirtilmemiş=30)")
    p.add_argument("--onek", default="B", help="Bulgu numarası öneki (B-01, BSCI-26-01 …)")
    p.add_argument("--gizle", nargs="*", default=[], metavar="AD", help="Raporda geçen ve gönderilmeyecek kişi adları")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "denetim_bulgulari.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    try:
        terminler = termin_oku(a.termin)
    except ValueError as h:
        print(f"[X] --termin {h}")
        return 2
    try:
        s = calistir(a.rapor, a.cikti, terminler, a.onek, a.gizle, a.tur, a.evet)
    except (llm.LLMHatasi, belge.BelgeHatasi) as h:
        print(f"[X] {h}")
        return 1
    for b in s["bulgular"]:
        print(f"[OK] {b['derece']} · {b['konu']} · termin {b['termin']:%d.%m.%Y}" if b["termin"] else f"[OK] {b['derece']} · {b['konu']}")
        for u in b["uyari"]:
            print(f"[!]   {u}")
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
