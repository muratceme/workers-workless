"""
Hasar İhbarı Ön Kayıt — Workers / Workless AI Agent
Hasar Yönetimi › Hasar Uzman Yardımcısı

1. Kod ihbar metninden plakaları (Türkiye plaka biçimi) ve poliçe listesinde geçen poliçe numaralarını bulur.
2. Model her ihbarı ön kayıt formatına çevirir: poliçe no, branş, hasar tarihi ve saati, il / ilçe / yer, hasar türü,
   tahmini tutar, sigortalı ve karşı taraf plakası, tutanak türü, yaralanma, kısa olay özeti, eksik bilgiler ve
   sigortalıdan istenecek belgeler. Her alan için metinden alıntı verir; metinde yoksa boş bırakır.
3. Kod denetler: poliçe no ve tutar metinde geçiyor mu, hasar tarihi metinde yazılı mı ya da "dün / bugün" gibi
   ihbar tarihinden çıkarılabiliyor mu, plakalar metindeki plakalardan mı, alıntılar metinde var mı. Sorun varsa
   modele geri gönderilir (en fazla --tur kez).
4. Poliçe listesi verilirse kod eşleştirir: poliçe bulunamadı, hasar tarihi poliçe süresi dışında, branş ya da plaka
   uyuşmuyor; poliçe no yazılmamışsa plakadan poliçe önerir. İhbar, hasar tarihinden sonra --ihbar-gun iş gününden
   (varsayılan 5) geç yapılmışsa uyarır. Yaralanma varsa ihbar öncelikli işaretlenir.

Kullanım:
    python agent.py                                           # örnek ihbarlar + poliçe listesi
    python agent.py --girdi ihbarlar.xlsx --policeler policeler.xlsx --ihbar-gun 5
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

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
PAKET = 10
PLAKA = re.compile(r"\b(0[1-9]|[1-7]\d|8[01])\s?([A-Z]{1,3})\s?(\d{2,4})\b")
BRANSLAR = ["Kasko", "Trafik", "Konut", "DASK", "İşyeri", "Nakliyat", "Mühendislik", "Sorumluluk", "Sağlık", "Ferdi Kaza", "Belirsiz"]
TURLER = ["Trafik kazası (çarpışma)", "Park hâlinde hasar", "Cam kırılması", "Su / tesisat hasarı", "Yangın", "Hırsızlık", "Doğal afet (sel, fırtına, deprem)",
          "Elektrik / elektronik hasar", "Bedeni hasar (yaralanma)", "Üçüncü şahıs sorumluluk", "Diğer"]
TUTANAK = ["Anlaşmalı kaza tespit tutanağı", "Polis / jandarma tutanağı", "İtfaiye raporu", "Yok", "Belirtilmemiş"]
ALANLAR = ["police_no", "hasar_tarihi", "yer", "tutar", "sigortali_plaka", "karsi_plaka", "tutanak", "yaralanma"]
SEMA = {
    "type": "object",
    "properties": {"ihbarlar": {"type": "array", "items": {
        "type": "object",
        "properties": {
            "id": S, "police_no": S, "brans": {"type": "string", "enum": BRANSLAR}, "hasar_tarihi": {"type": "string", "description": "GG.AA.YYYY ya da boş"},
            "hasar_saati": S, "il": S, "ilce": S, "yer": S, "hasar_turu": {"type": "string", "enum": TURLER},
            "tahmini_tutar": {"type": ["number", "null"]}, "sigortali_plaka": S, "karsi_plaka": S, "tutanak": {"type": "string", "enum": TUTANAK},
            "yaralanma": {"type": "boolean"}, "olay_ozeti": S, "eksik_bilgiler": {"type": "array", "items": S}, "istenecek_belgeler": {"type": "array", "items": S},
            "alintilar": {"type": "array", "items": {"type": "object", "properties": {"alan": {"type": "string", "enum": ALANLAR}, "alinti": S},
                                                     "required": ["alan", "alinti"], "additionalProperties": False}},
        },
        "required": ["id", "police_no", "brans", "hasar_tarihi", "hasar_saati", "il", "ilce", "yer", "hasar_turu", "tahmini_tutar", "sigortali_plaka",
                     "karsi_plaka", "tutanak", "yaralanma", "olay_ozeti", "eksik_bilgiler", "istenecek_belgeler", "alintilar"],
        "additionalProperties": False}}},
    "required": ["ihbarlar"],
    "additionalProperties": False,
}
YARALANMA = re.compile(r"(yaral|hastane|ambulans|vefat|yoğun bakım|acil servis)", re.I)
YARALANMA_YOK = re.compile(r"(yaralanan yok|yaralanma yok|yaralı yok|kimse yaralanmadı)", re.I)


def plaka_norm(p: str) -> str:
    m = PLAKA.search(str(p or "").upper().replace("İ", "I"))
    return f"{m.group(1)} {m.group(2)} {m.group(3)}" if m else ""


def plakalar(metin: str) -> list[str]:
    return list(dict.fromkeys(f"{a} {b} {c}" for a, b, c in PLAKA.findall(str(metin).upper().replace("İ", "I"))))


def is_gunu_farki(bas: date, bit: date) -> int:
    """bas'tan sonraki günden bit'e kadar (dahil) hafta içi gün sayısı. Resmî tatiller hesaba katılmaz."""
    return sum(1 for i in range(1, (bit - bas).days + 1) if (bas + timedelta(days=i)).weekday() < 5)


def tarih_dayanagi(tarih: date, metin: str, ihbar: date | None) -> bool:
    k = katla(metin)
    yazimlar = {tarih.strftime(f) for f in ("%d.%m.%Y", "%d/%m/%Y", "%d.%m.%y", "%Y-%m-%d", "%d %m %Y")}
    yazimlar |= {f"{tarih.day} {tarih.month} {tarih.year}", f"{tarih.day} {tarih.month}"}
    if any(katla(y) in k for y in yazimlar):
        return True
    if ihbar:
        if re.search(r"\bdun\b", k) and tarih == ihbar - timedelta(days=1):
            return True
        if re.search(r"\b(bugun|bu sabah|bu gece|az once)\b", k) and tarih == ihbar:
            return True
        if re.search(r"\b(evvelsi gun|onceki gun)\b", k) and tarih == ihbar - timedelta(days=2):
            return True
    return False


def ihbarlari_oku(yol: Path) -> list[dict]:
    r = kayitlar(yol, {"id": ("İhbar No", "Kayıt No", "No"), "zaman": ("İhbar Tarihi", "Tarih", "Geliş Tarihi"), "kanal": ("Kanal", "Kaynak"),
                       "metin": ("Metin", "İhbar Metni", "Mesaj", "Açıklama", "E-posta")}, ("metin",))
    return [{"id": str(x["id"] or f"IH{i}"), "zaman": zaman_coz(x["zaman"]), "kanal": str(x["kanal"]), "metin": str(x["metin"])}
            for i, x in enumerate(r, 1) if str(x["metin"]).strip()]


def policeleri_oku(yol: Path) -> dict[str, dict]:
    r = kayitlar(yol, {"no": ("Poliçe No", "Police No", "Poliçe"), "sigortali": ("Sigortalı", "Sigorta Ettiren"), "brans": ("Branş", "Ürün"),
                       "bas": ("Başlangıç", "Başlangıç Tarihi"), "bit": ("Bitiş", "Bitiş Tarihi"), "plaka": ("Plaka",)}, ("no",))
    return {str(x["no"]).strip().upper(): {**x, "bas": tarih_coz(x["bas"]), "bit": tarih_coz(x["bit"]), "plaka": plaka_norm(x["plaka"])} for x in r if x["no"]}


def denetle(ihbarlar: list[dict], sonuc: dict[str, dict]) -> list[str]:
    sorun = []
    for i in ihbarlar:
        s = sonuc.get(i["id"])
        if not s:
            sorun.append(f"{i['id']}: sonuç verilmedi")
            continue
        m, ad = i["metin"], i["id"]
        if s["police_no"] and not metinde_var(s["police_no"], m):
            sorun.append(f"{ad}: poliçe no '{s['police_no']}' metinde geçmiyor (metinde yoksa boş bırakın)")
        if s["tahmini_tutar"] is not None and sayi_yaz(s["tahmini_tutar"]) not in sayilar(m):
            sorun.append(f"{ad}: tahmini tutar {sayi_yaz(s['tahmini_tutar'])} metinde geçmiyor (tutar yazılmamışsa null)")
        if s["hasar_tarihi"]:
            t = tarih_coz(s["hasar_tarihi"])
            if not t:
                sorun.append(f"{ad}: hasar tarihi '{s['hasar_tarihi']}' GG.AA.YYYY biçiminde değil")
            elif not tarih_dayanagi(t, m, i["zaman"].date() if i["zaman"] else None):
                sorun.append(f"{ad}: hasar tarihi {s['hasar_tarihi']} metinde yazmıyor ve ihbar tarihinden (dün / bugün) çıkarılamıyor; tahmin etmeyin, boş bırakın")
            elif i["zaman"] and t > i["zaman"].date():
                sorun.append(f"{ad}: hasar tarihi {s['hasar_tarihi']} ihbar tarihinden sonra")
        bulunan = plakalar(m)
        for alan, ad2 in (("sigortali_plaka", "sigortalı plakası"), ("karsi_plaka", "karşı taraf plakası")):
            if s[alan] and plaka_norm(s[alan]) not in bulunan:
                sorun.append(f"{ad}: {ad2} '{s[alan]}' metindeki plakalardan değil ({', '.join(bulunan) or 'metinde plaka yok'})")
        if s["sigortali_plaka"] and s["karsi_plaka"] and plaka_norm(s["sigortali_plaka"]) == plaka_norm(s["karsi_plaka"]):
            sorun.append(f"{ad}: sigortalı ve karşı taraf plakası aynı")
        yar = bool(YARALANMA.search(m)) and not YARALANMA_YOK.search(m)
        if yar and not s["yaralanma"]:
            sorun.append(f"{ad}: metinde yaralanma / hastane ifadesi var ama yaralanma false")
        for a in s["alintilar"]:
            if not metinde_var(a["alinti"], m):
                sorun.append(f"{ad}: '{a['alan']}' alıntısı metinde yok: '{a['alinti'][:50]}'")
    return sorun


def police_kontrol(i: dict, s: dict, pol: dict[str, dict] | None, ihbar_gun: int) -> tuple[list[str], str]:
    """(uyarılar, önerilen poliçe)"""
    u, oneri = [], ""
    t = tarih_coz(s["hasar_tarihi"])
    if s["yaralanma"]:
        u.append("ÖNCELİKLİ: bedeni hasar / yaralanma bildirildi")
    if t and i["zaman"]:
        g = is_gunu_farki(t, i["zaman"].date())
        if g > ihbar_gun:
            u.append(f"İhbar hasardan {g} iş günü sonra yapılmış (sınır {ihbar_gun}); gecikme nedeni sorulmalı")
    if not s["hasar_tarihi"]:
        u.append("Hasar tarihi yok")
    if pol is None:
        return u, oneri
    p = pol.get(s["police_no"].strip().upper()) if s["police_no"] else None
    if s["police_no"] and not p:
        u.append(f"Poliçe {s['police_no']} poliçe listesinde bulunamadı")
    if not s["police_no"]:
        adaylar = [no for no, x in pol.items() if x["plaka"] and x["plaka"] in plakalar(i["metin"])]
        if adaylar:
            oneri = ", ".join(adaylar)
            u.append(f"Poliçe no yazılmamış; plakadan önerilen poliçe: {oneri} (sigortalıyla teyit edin)")
            p = pol[adaylar[0]] if len(adaylar) == 1 else None
        else:
            u.append("Poliçe no yazılmamış ve plakadan eşleşme bulunamadı")
    if p:
        if t and p["bas"] and p["bit"] and not (p["bas"] <= t <= p["bit"]):
            u.append(f"Hasar tarihi {t:%d.%m.%Y} poliçe süresi dışında ({p['bas']:%d.%m.%Y}–{p['bit']:%d.%m.%Y})")
        if p["brans"] and s["brans"] != "Belirsiz" and katla(p["brans"]) != katla(s["brans"]):
            u.append(f"Branş uyuşmuyor: ihbarda {s['brans']}, poliçede {p['brans']}")
        if p["plaka"] and s["sigortali_plaka"] and plaka_norm(s["sigortali_plaka"]) != p["plaka"]:
            u.append(f"Plaka uyuşmuyor: ihbarda {plaka_norm(s['sigortali_plaka'])}, poliçede {p['plaka']}")
    return u, oneri


def uret(ihbarlar: list[dict], duzeltme=None, onceki=None) -> dict[str, dict]:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    sonuc = {}
    if duzeltme is not None:  # düzeltmede yalnız sorunlu kayıtlar yeniden sorulur
        sorunlu = {s.split(":")[0] for s in duzeltme}
        sonuc.update({a: b for a, b in (onceki or {}).items() if a not in sorunlu})
        ihbarlar = [x for x in ihbarlar if x["id"] in sorunlu]
    for j in range(0, len(ihbarlar), PAKET):
        parca = ihbarlar[j:j + PAKET]
        kimlikler = {i["id"] for i in parca}
        satirlar = [f'<ihbar id="{i["id"]}" ihbar_tarihi="{i["zaman"]:%d.%m.%Y %H:%M}" kanal="{i["kanal"]}" metindeki_plakalar="{", ".join(plakalar(i["metin"])) or "-"}">\n'
                    f"{llm.maskele(i['metin'])}\n</ihbar>" if i["zaman"] else
                    f'<ihbar id="{i["id"]}" kanal="{i["kanal"]}" metindeki_plakalar="{", ".join(plakalar(i["metin"])) or "-"}">\n{llm.maskele(i["metin"])}\n</ihbar>'
                    for i in parca]
        d = [x for x in (duzeltme or []) if x.split(":")[0] in kimlikler]
        if duzeltme is not None and not d:
            sonuc.update({k: v for k, v in (onceki or {}).items() if k in kimlikler})
            continue
        o = {k: v for k, v in (onceki or {}).items() if k in kimlikler}
        mesaj = "\n".join(["<ihbarlar>", *satirlar, "</ihbarlar>"]) + duzeltme_blogu(d, list(o.values()))
        for s in llm.json_iste(sistem, mesaj, SEMA).get("ihbarlar", []):
            if s.get("id") in kimlikler:
                sonuc[s["id"]] = s
    return sonuc


def rapor_yaz(cikti: Path, ihbarlar, sonuc, uyarilar, oneriler, kalan, tur) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ön Kayıt"
    baslik(ws, ["İhbar No", "İhbar Tarihi", "Kanal", "Öncelik", "Poliçe No", "Önerilen Poliçe", "Branş", "Hasar Tarihi", "Saat", "İl", "İlçe", "Yer",
                "Hasar Türü", "Tahmini Tutar (TL)", "Sigortalı Plaka", "Karşı Plaka", "Tutanak", "Yaralanma", "Olay Özeti", "Eksik Bilgiler",
                "İstenecek Belgeler", "Kontrol Uyarıları", "Onay"],
           (12, 15, 11, 10, 16, 16, 10, 11, 7, 10, 12, 26, 22, 12, 12, 12, 22, 9, 46, 36, 40, 50, 8))
    for i in sorted(ihbarlar, key=lambda x: (not sonuc.get(x["id"], {}).get("yaralanma"), x["zaman"] or datetime.min)):
        s = sonuc.get(i["id"])
        if not s:
            ws.append([i["id"], i["zaman"], i["kanal"], "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "AI yanıt vermedi"])
            continue
        u = uyarilar[i["id"]]
        ws.append([i["id"], i["zaman"], i["kanal"], "Öncelikli" if s["yaralanma"] else "Normal", s["police_no"], oneriler[i["id"]], s["brans"],
                   s["hasar_tarihi"], s["hasar_saati"], s["il"], s["ilce"], s["yer"], s["hasar_turu"], s["tahmini_tutar"], s["sigortali_plaka"],
                   s["karsi_plaka"], s["tutanak"], "Evet" if s["yaralanma"] else "Hayır", s["olay_ozeti"], "; ".join(s["eksik_bilgiler"]),
                   "; ".join(s["istenecek_belgeler"]), "; ".join(u), ""])
        ws.cell(ws.max_row, 2).number_format = "DD.MM.YYYY HH:MM"
        ws.cell(ws.max_row, 14).number_format = "#,##0"
        boya(ws, range(5, 22), AI_DOLGU)
        if s["yaralanma"]:
            boya(ws, [4], KIRMIZI)
        if u:
            boya(ws, [22], SARI)
        boya(ws, [23], SARI)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:W{ws.max_row}"
    son_not(ws, "Önemli: Yapay zekâ ön kaydıdır. Hasar dosyası açılmadan önce ihbar metni ve poliçe bilgileriyle karşılaştırılmalıdır.", kalan, tur)
    hizala(ws)
    m = wb.create_sheet("İhbar Metinleri")
    baslik(m, ["İhbar No", "Metin", "Metindeki Plakalar"], (12, 110, 24))
    for i in ihbarlar:
        m.append([i["id"], i["metin"], ", ".join(plakalar(i["metin"]))])
    hizala(m)
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(girdi: Path, cikti: Path, police_yolu: Path | None = None, ihbar_gun: int = 5, tur: int = 2, evet: bool = False) -> dict:
    ihbarlar = ihbarlari_oku(girdi)
    if not ihbarlar:
        raise llm.LLMHatasi(f"{girdi.name}: ihbar yok.")
    pol = policeleri_oku(police_yolu) if police_yolu else None
    print(f"[OK] {len(ihbarlar)} ihbar" + (f" · {len(pol)} poliçe" if pol is not None else ""))
    llm.onay_al(f"{len(ihbarlar)} ihbar metni (telefon / e-posta / TCKN / IBAN maskeli; ad ve plaka maskelenmez) gönderilecek.", evet)
    sonuc, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(ihbarlar, d, o), lambda x: denetle(ihbarlar, x), tur)
    uyarilar, oneriler = {}, {}
    for i in ihbarlar:
        if i["id"] in sonuc:
            uyarilar[i["id"]], oneriler[i["id"]] = police_kontrol(i, sonuc[i["id"]], pol, ihbar_gun)
    rapor_yaz(cikti, ihbarlar, sonuc, uyarilar, oneriler, kalan, yapilan)
    return {"ihbarlar": ihbarlar, "sonuc": sonuc, "uyarilar": uyarilar, "oneriler": oneriler, "kalan": kalan, "tur": yapilan}


def main(argv: list[str] | None = None) -> int:
    ortam_hazirla()
    p = argparse.ArgumentParser(description="Hasar ihbar e-posta ve formlarını ön kayıt formatına çevirir, poliçeyle karşılaştırır.")
    p.add_argument("--girdi", type=Path, default=BURASI / "ornek_veri" / "ihbarlar.csv",
                   help="İhbarlar (.xlsx/.csv): Metin; İhbar No, İhbar Tarihi, Kanal isteğe bağlı")
    p.add_argument("--policeler", type=Path, help="İsteğe bağlı poliçe listesi: Poliçe No, Sigortalı, Branş, Başlangıç, Bitiş, Plaka")
    p.add_argument("--ihbar-gun", type=int, default=5, help="Hasardan sonra ihbar için iş günü sınırı (varsayılan 5; poliçe şartlarına göre ayarlayın)")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "hasar_on_kayit.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    if a.girdi == BURASI / "ornek_veri" / "ihbarlar.csv" and not a.policeler:
        a.policeler = BURASI / "ornek_veri" / "policeler.csv"
    try:
        s = calistir(a.girdi, a.cikti, a.policeler, a.ihbar_gun, a.tur, a.evet)
    except llm.LLMHatasi as h:
        print(f"[X] {h}")
        return 1
    print(f"[OK] {len(s['sonuc'])}/{len(s['ihbarlar'])} ihbar ön kayda çevrildi · öncelikli: {sum(x['yaralanma'] for x in s['sonuc'].values())}")
    for k, u in s["uyarilar"].items():
        for x in u:
            print(f"[!] {k}: {x}")
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
