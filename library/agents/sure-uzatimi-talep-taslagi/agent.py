"""
Süre Uzatımı Talep Taslağı — Workers / Workless AI Agent
İnşaat Teknik Ofis › Teknik Ofis Mühendisi (Teknik Ofis Şefi ve Planlama Mühendisi de kullanır)

Günlük şantiye raporlarındaki gecikme olaylarından gün hesabını kodla yapar ve idareye süre uzatımı talep yazısı taslağı
hazırlar.
1. Kod günlük raporları okur (her rapor "Tarih: GG.AA.YYYY" satırıyla başlar); telefon, e-posta, kimlik numarası maskelenir.
2. Model her rapordaki gecikme olaylarını sınıflar: neden türü (hava, idare kaynaklı, mücbir sebep, yer teslimi / engel,
   malzeme / tedarik, yüklenici kaynaklı, İSG, diğer), çalışmanın durma derecesi (tam / kısmi / yok), etkilenen iş ve
   rapordan aynen alıntı.
3. Kod sınıflamayı denetler (alıntı raporda aynen geçiyor mu); sonra gün hesabını yapar: her gün bir kez sayılır;
   talep edilebilir nedenler (--talep-edilebilir) ile durmuş günler; yalnız yüklenici kaynaklı günler hariç; aynı gün
   yüklenici kaynaklı olayla çakışan günler "eşzamanlı gecikme" olarak işaretlenir. --basvuru-gun verilirse ilk olaydan
   son başvuru tarihi hesaplanır.
4. Model kodun özet tablosu ve verdiğiniz sözleşme maddeleriyle yazının gövdesini yazar; kod yazıdaki her sayı ve tarihin
   özette ya da raporlarda, her madde numarasının sözleşme metninde olduğunu denetler.

Kullanım:
    python agent.py                                                         # örnek 5 günlük rapor
    python agent.py --raporlar ./gunluk_raporlar --sozlesme maddeler.txt --is "B Blok kaba inşaat" --idare "Kurgu Belediyesi" --basvuru-gun 20
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
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
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
SARI = PatternFill("solid", fgColor="FFF4CE")
YESIL = PatternFill("solid", fgColor="E3F5E1")
UST = Alignment(vertical="top", wrap_text=True)
S = {"type": "string"}
TURLER = ["Hava muhalefeti", "İdare kaynaklı", "Mücbir sebep", "Yer teslimi / engel", "Malzeme / tedarik", "Yüklenici kaynaklı", "İSG / kaza", "Diğer"]
VARSAYILAN_TALEP = ["Hava muhalefeti", "İdare kaynaklı", "Mücbir sebep", "Yer teslimi / engel"]
MADDE = re.compile(r"\b(?:madde|md\.?)\s*(\d+(?:\.\d+)*)", re.I)


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


def belirtecler(metin: str) -> set[str]:
    """Metindeki tarih ve sayılar (madde numaraları hariç), karşılaştırma için."""
    m = MADDE.sub(" ", str(metin))
    m = re.sub(r"(?m)^\s*\d{1,2}[.)]\s", " ", m)                         # satır başı liste numaraları ("1. ", "2) ")
    out = set(re.findall(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b", m))
    m = re.sub(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b|\b\d{1,2}:\d{2}\b", " ", m)
    out |= {f"{v.normalize():f}" for x in re.findall(r"\d+(?:[.,]\d+)*(?:/\d+)?", m) for v in [para(x.split("/")[0])] if v is not None}
    out |= set(re.findall(r"\b\d{4}/\d+\b", str(metin)))
    return out


def tarih_coz(x) -> date | None:
    m = re.search(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4})\b", str(x or ""))
    if not m:
        return None
    try:
        return date(int(m[3]), int(m[2]), int(m[1]))
    except ValueError:
        return None


# ----------------------------------------------------------------------------
# Girdi
# ----------------------------------------------------------------------------

def raporlari_oku(klasor: Path) -> list[dict]:
    if klasor.is_dir():
        dosyalar = sorted(p for p in klasor.iterdir() if p.is_file() and p.suffix.lower() in belge.DESTEKLENEN)
    elif klasor.exists():
        dosyalar = [klasor]
    else:
        raise belge.BelgeHatasi(f"{klasor}: bulunamadı.")
    out = []
    for p in dosyalar:
        metin = belge.metin_oku(p)
        m = re.search(r"^\s*Tarih\s*:\s*(\d{1,2}\.\d{1,2}\.\d{4})", metin, re.I | re.M)
        t = tarih_coz(m[1]) if m else None
        if t is None:
            raise llm.LLMHatasi(f"{p.name}: 'Tarih: GG.AA.YYYY' satırı yok.")
        out.append({"dosya": p.name, "tarih": t, "metin": llm.maskele(metin)})
    if not out:
        raise llm.LLMHatasi(f"{klasor}: günlük rapor yok.")
    out.sort(key=lambda r: r["tarih"])
    tekrar = [r["tarih"] for i, r in enumerate(out) if i and out[i - 1]["tarih"] == r["tarih"]]
    if tekrar:
        raise llm.LLMHatasi(f"Aynı tarihli birden fazla rapor: {', '.join(f'{t:%d.%m.%Y}' for t in tekrar)} (tek dosyada birleştirin).")
    for i, r in enumerate(out, 1):
        r["no"] = f"R{i}"
    return out


# ----------------------------------------------------------------------------
# Model 1: olay sınıflama
# ----------------------------------------------------------------------------

OLAY = {"type": "object", "properties": {"tur": {"type": "string", "enum": TURLER}, "durma": {"type": "string", "enum": ["Tam", "Kısmi", "Yok"]},
                                         "etkilenen_is": S, "aciklama": S, "alinti": S},
        "required": ["tur", "durma", "etkilenen_is", "aciklama", "alinti"], "additionalProperties": False}
SEMA_OLAY = {"type": "object", "properties": {"olaylar": {"type": "array", "items": OLAY}}, "required": ["olaylar"], "additionalProperties": False}


def denetle_olay(r: dict, c: dict) -> list[str]:
    sorun = []
    for i, o in enumerate(c["olaylar"], 1):
        if not metinde_var(o["alinti"], r["metin"]):
            sorun.append(f"{r['no']}: olay {i}: alıntı '{o['alinti'][:50]}' raporda aynen geçmiyor")
        if not o["aciklama"].strip():
            sorun.append(f"{r['no']}: olay {i}: açıklama boş")
    return sorun


def uret_olay(raporlar: list[dict], duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    sonuc = dict(onceki or {})
    hedef = raporlar if duzeltme is None else [r for r in raporlar if any(s.startswith(r["no"] + ":") for s in duzeltme)]
    for r in hedef:
        mesaj = f"<rapor tarih=\"{r['tarih']:%d.%m.%Y}\">\n{r['metin']}\n</rapor>"
        d = [s.split(": ", 1)[1] for s in (duzeltme or []) if s.startswith(r["no"] + ":")]
        if d:
            mesaj += ("\n<duzeltme>\nÖnceki sınıflama: " + json.dumps(sonuc.get(r["no"]), ensure_ascii=False) + "\nKodun bulduğu sorunlar:\n"
                      + "\n".join(f"- {s}" for s in d) + "\n</duzeltme>")
        sonuc[r["no"]] = llm.json_iste(sistem, mesaj, SEMA_OLAY)
    return sonuc


def gun_hesabi(raporlar: list[dict], olaylar: dict, talep_turleri: list[str], basvuru_gun: int | None) -> dict:
    """Kod: gün bazında durum, neden bazında gün sayısı, talep edilebilir gün, eşzamanlı gecikme ve başvuru tarihi."""
    gunler = []
    for r in raporlar:
        ol = [o for o in olaylar[r["no"]]["olaylar"] if o["durma"] != "Yok"]
        talep = [o for o in ol if o["tur"] in talep_turleri]
        yuk = [o for o in ol if o["tur"] == "Yüklenici kaynaklı" or o["tur"] == "Malzeme / tedarik"]
        if not ol:
            continue
        durum = "Tam" if any(o["durma"] == "Tam" for o in talep) else "Kısmi" if talep else "—"
        gunler.append({"tarih": r["tarih"], "talep": bool(talep), "durum": durum, "nedenler": sorted({o["tur"] for o in talep}),
                       "eszamanli": bool(talep and yuk), "haric": not talep, "olaylar": ol})
    neden = defaultdict(lambda: {"Tam": 0, "Kısmi": 0})
    for g in gunler:
        for n in g["nedenler"]:
            neden[n]["Tam" if any(o["tur"] == n and o["durma"] == "Tam" for o in g["olaylar"]) else "Kısmi"] += 1
    talep_gunleri = [g for g in gunler if g["talep"]]
    ilk = talep_gunleri[0]["tarih"] if talep_gunleri else None
    return {"gunler": gunler, "neden": dict(neden), "toplam": len(talep_gunleri), "tam": sum(g["durum"] == "Tam" for g in talep_gunleri),
            "kismi": sum(g["durum"] == "Kısmi" for g in talep_gunleri), "eszamanli": [g["tarih"] for g in gunler if g["eszamanli"]],
            "haric": [g["tarih"] for g in gunler if g["haric"]], "ilk": ilk, "son": talep_gunleri[-1]["tarih"] if talep_gunleri else None,
            "basvuru_son": ilk + timedelta(days=basvuru_gun) if ilk and basvuru_gun else None}


def ozet_metni(h: dict) -> str:
    l = [f"Talep edilebilir gün: {h['toplam']} (tam duruş {h['tam']}, kısmi duruş {h['kismi']})",
         f"Dönem: {h['ilk']:%d.%m.%Y} – {h['son']:%d.%m.%Y}" if h["ilk"] else "Dönem: -"]
    l += [f"- {n}: {v['Tam']} gün tam, {v['Kısmi']} gün kısmi" for n, v in h["neden"].items()]
    l.append("Gün listesi:")
    for g in h["gunler"]:
        if g["talep"]:
            ek = " · aynı gün yüklenici kaynaklı olay da var (eşzamanlı gecikme)" if g["eszamanli"] else ""
            l.append(f"- {g['tarih']:%d.%m.%Y}: {g['durum']} duruş · {', '.join(g['nedenler'])} · etkilenen: "
                     f"{'; '.join(o['etkilenen_is'] for o in g['olaylar'] if o['tur'] in g['nedenler'])}{ek}")
    if h["haric"]:
        l.append("Talep dışı bırakılan günler (yalnız yüklenici / tedarik kaynaklı): " + ", ".join(f"{t:%d.%m.%Y}" for t in h["haric"]))
    if h["basvuru_son"]:
        l.append(f"Son başvuru tarihi (ilk olay + başvuru süresi): {h['basvuru_son']:%d.%m.%Y}")
    return "\n".join(l)


# ----------------------------------------------------------------------------
# Model 2: yazı
# ----------------------------------------------------------------------------

SEMA_YAZI = {"type": "object", "properties": {"konu": S, "govde": S, "dayanak_maddeler": {"type": "array", "items": S}, "ekler": {"type": "array", "items": S}},
             "required": ["konu", "govde", "dayanak_maddeler", "ekler"], "additionalProperties": False}


def denetle_yazi(c: dict, h: dict, raporlar: list[dict], sozlesme: str) -> list[str]:
    sorun, metin = [], c["konu"] + "\n" + c["govde"]
    izinli = belirtecler(ozet_metni(h)) | belirtecler("\n".join(r["metin"] for r in raporlar))
    fazla = sorted(belirtecler(metin) - izinli)
    if fazla:
        sorun.append(f"Yazı: özette ya da raporlarda olmayan sayı / tarih: {', '.join(fazla)}")
    gecerli = {g.lower() for g in MADDE.findall(sozlesme)}
    yok = sorted({g.lower() for g in MADDE.findall(metin + " " + " ".join(c["dayanak_maddeler"]))} - gecerli)
    if yok:
        sorun.append(f"Yazı: sözleşme metninde olmayan madde: {', '.join(yok)}")
    if str(h["toplam"]) not in re.findall(r"\b\d+\b", c["govde"]):
        sorun.append(f"Yazı: talep edilen toplam gün ({h['toplam']}) gövdede yok")
    if re.search(r"\b(kesinlikle|mutlaka|hukuki yollara|dava|tazminat talep)\w*", metin, re.I):
        sorun.append("Yazı: kesinlik / tehdit ifadesi; yazı nesnel olmalı")
    return sorun


def uret_yazi(h: dict, sozlesme: str, is_adi: str, idare: str, duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt_yazi.md").read_text(encoding="utf-8")
    mesaj = (f"<is>{is_adi or 'belirtilmedi'}</is>\n<idare>{idare or 'belirtilmedi'}</idare>\n<sozlesme>\n{sozlesme or '(verilmedi)'}\n</sozlesme>\n"
             f"<ozet>\n{ozet_metni(h)}\n</ozet>")
    if duzeltme:
        mesaj += ("\n<duzeltme>\nÖnceki yazı: " + json.dumps(onceki, ensure_ascii=False) + "\nKodun bulduğu sorunlar:\n"
                  + "\n".join(f"- {s}" for s in duzeltme) + "\n</duzeltme>")
    return llm.json_iste(sistem, mesaj, SEMA_YAZI)


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

def _baslik(ws, basliklar, genislik) -> None:
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font, h.alignment = BASLIK_DOLGU, BASLIK_YAZI, UST
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def rapor_yaz(cikti: Path, raporlar, olaylar, h, yazi, kalan, tur, imza) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Olaylar"
    _baslik(ws, ["Tarih", "Rapor", "Neden Türü", "Durma", "Etkilenen İş", "Açıklama", "Rapordan Alıntı"], (11, 16, 18, 8, 30, 50, 60))
    for r in raporlar:
        for o in olaylar[r["no"]]["olaylar"]:
            ws.append([r["tarih"], r["dosya"], o["tur"], o["durma"], o["etkilenen_is"], o["aciklama"], o["alinti"]])
            ws.cell(ws.max_row, 1).number_format = "DD.MM.YYYY"
            for j in (3, 4, 5, 6):
                ws.cell(ws.max_row, j).fill = AI_DOLGU
            for x in ws[ws.max_row]:
                x.alignment = UST
    g = wb.create_sheet("Gün Hesabı")
    _baslik(g, ["Tarih", "Talep Edilebilir", "Duruş", "Nedenler", "Eşzamanlı Gecikme", "Not"], (11, 10, 8, 40, 10, 60))
    for x in h["gunler"]:
        g.append([x["tarih"], "Evet" if x["talep"] else "Hayır", x["durum"], ", ".join(x["nedenler"]) or ", ".join(sorted({o["tur"] for o in x["olaylar"]})),
                  "Evet" if x["eszamanli"] else "", "Yalnız yüklenici / tedarik kaynaklı: talep dışı" if x["haric"] else
                  ("Aynı gün yüklenici kaynaklı olay var: idare itiraz edebilir" if x["eszamanli"] else "")])
        g.cell(g.max_row, 1).number_format = "DD.MM.YYYY"
        if x["eszamanli"]:
            g.cell(g.max_row, 5).fill = SARI
    g.append([])
    g.append([f"Toplam talep edilebilir gün: {h['toplam']} (tam {h['tam']}, kısmi {h['kismi']})"])
    if h["basvuru_son"]:
        g.append([f"Son başvuru tarihi: {h['basvuru_son']:%d.%m.%Y}"])
    y = wb.create_sheet("Yazı")
    y.column_dimensions["A"].width = 110
    for s in (f"Konu: {yazi['konu']}", "", *yazi["govde"].splitlines(), "", "Ekler:", *[f"  {i}. {e}" for i, e in enumerate(yazi["ekler"], 1)], "", "Saygılarımızla,", imza):
        y.append([s])
        y.cell(y.max_row, 1).alignment = UST
    k = wb.create_sheet("Kontrol")
    k.append([f"Düzeltme turu: {tur} · Model: {llm.kullanim_ozeti()} · {datetime.now():%d.%m.%Y %H:%M}"])
    k.append(["Son kontrol: " + ("; ".join(kalan) if kalan else "Kod denetiminden sorunsuz geçti")])
    k.cell(2, 1).fill = KIRMIZI if kalan else YESIL
    k.append(["Önemli: Gün hesabı kodla yapılmıştır; neden sınıflaması yapay zekâ önerisidir (mor) ve teknik ofis tarafından kontrol edilmelidir. "
              "Süre uzatımı hakkı ve başvuru süresi sözleşmenize göre değişir; yazıyı göndermeden önce sözleşme ve hukuk birimiyle kontrol edin."])
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)
    txt = cikti.with_suffix(".txt")
    ekler = "\n".join(f"{i}. {e}" for i, e in enumerate(yazi["ekler"], 1))
    txt.write_text(f"Konu: {yazi['konu']}\n\n{yazi['govde']}\n\nEkler:\n{ekler}\n\nSaygılarımızla,\n{imza}\n", encoding="utf-8")
    return txt


def calistir(girdi: Path, cikti: Path, sozlesme_yolu: Path | None = None, is_adi: str = "", idare: str = "", talep_turleri: list[str] | None = None,
             basvuru_gun: int | None = None, imza: str = "Yüklenici Teknik Ofis", tur: int = 2, evet: bool = False) -> dict:
    raporlar = raporlari_oku(girdi)
    if sozlesme_yolu is not None and not sozlesme_yolu.exists():
        raise llm.LLMHatasi(f"{sozlesme_yolu}: sözleşme maddeleri dosyası bulunamadı.")
    sozlesme = belge.metin_oku(sozlesme_yolu) if sozlesme_yolu else ""
    talep_turleri = talep_turleri or VARSAYILAN_TALEP
    print(f"[OK] {len(raporlar)} günlük rapor ({raporlar[0]['tarih']:%d.%m.%Y} – {raporlar[-1]['tarih']:%d.%m.%Y})")
    llm.onay_al(f"{len(raporlar)} günlük raporun metni ve sözleşme maddeleri gönderilecek (telefon, e-posta, TCKN maskeli).", evet)
    olaylar, kalan1, t1 = duzeltme_dongusu(lambda d, o: uret_olay(raporlar, d, o), lambda x: [s for r in raporlar for s in denetle_olay(r, x[r["no"]])], tur)
    h = gun_hesabi(raporlar, olaylar, talep_turleri, basvuru_gun)
    print(f"[OK] Talep edilebilir gün: {h['toplam']} (tam {h['tam']}, kısmi {h['kismi']}) · eşzamanlı: {len(h['eszamanli'])}")
    if h["toplam"] == 0:
        yazi, kalan2, t2 = {"konu": "", "govde": "Talep edilebilir gecikme günü bulunamadı; yazı hazırlanmadı.", "dayanak_maddeler": [], "ekler": []}, [], 0
    else:
        yazi, kalan2, t2 = duzeltme_dongusu(lambda d, o: uret_yazi(h, sozlesme, is_adi, idare, d, o), lambda x: denetle_yazi(x, h, raporlar, sozlesme), tur)
    txt = rapor_yaz(cikti, raporlar, olaylar, h, yazi, kalan1 + kalan2, t1 + t2, imza)
    return {"raporlar": raporlar, "olaylar": olaylar, "hesap": h, "yazi": yazi, "kalan": kalan1 + kalan2, "tur": t1 + t2, "txt": txt}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    p = argparse.ArgumentParser(description="Günlük şantiye raporlarından gecikme günlerini hesaplar ve süre uzatımı talep yazısı taslağı hazırlar.")
    p.add_argument("--raporlar", type=Path, default=ORNEK / "raporlar", help="Günlük raporlar klasörü (.txt/.docx/.pdf); her rapor 'Tarih: GG.AA.YYYY' satırı içerir")
    p.add_argument("--sozlesme", type=Path, default=None, help="Sözleşme / şartname maddeleri metni")
    p.add_argument("--is", dest="is_adi", default="", help="İşin adı")
    p.add_argument("--idare", default="", help="İdarenin adı")
    p.add_argument("--talep-edilebilir", nargs="*", default=None, choices=TURLER, metavar="TÜR", help="Talep edilebilir neden türleri")
    p.add_argument("--basvuru-gun", type=int, default=None, help="Sözleşmedeki başvuru süresi (gün); ilk olaydan son başvuru tarihi hesaplanır")
    p.add_argument("--imza", default="Yüklenici Teknik Ofis")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "sure_uzatimi.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    ornek = a.raporlar == ORNEK / "raporlar"
    try:
        s = calistir(a.raporlar, a.cikti, a.sozlesme or (ORNEK / "sozlesme_maddeleri.txt" if ornek else None),
                     a.is_adi or ("B Blok kaba inşaat işleri" if ornek else ""), a.idare or ("Kurgu Belediyesi Fen İşleri Müdürlüğü" if ornek else ""),
                     a.talep_edilebilir, a.basvuru_gun if a.basvuru_gun is not None else (20 if ornek else None), a.imza, a.tur, a.evet)
    except (llm.LLMHatasi, belge.BelgeHatasi) as h:
        print(f"[X] {h}")
        return 1
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    print(f"[OK] Rapor: {a.cikti.resolve()} (+ {s['txt'].name})")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
