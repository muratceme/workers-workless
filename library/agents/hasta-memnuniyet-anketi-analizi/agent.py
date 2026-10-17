"""
Hasta Memnuniyet Anketi Analizi — Gerisi Bende AI Agent
Hasta Hizmetleri › Hasta İlişkileri ve Hasta Hakları Yöneticisi

1. Kod anket puanlarını (1–5) hesaplar: birim ve hekim bazında yanıt sayısı, soru ortalamaları, genel ortalama ve
   memnuniyet oranı (4–5 verilen yanıtların payı). --min-n'den (varsayılan 10) az yanıtı olan hekim "yetersiz örnek"
   sayılır, sıralamaya girmez. Ortalaması --hedef'in (varsayılan 4,0) altında kalan birim × soru hücreleri "düşük" olur.
2. Model yalnız yorumları görür (hekim adı ve hasta bilgisi gönderilmez; yorumdaki hekim adları [HEKİM] yapılır):
   her yorumu temaya ve duyguya ayırır, birim bazında iyileştirme başlıkları önerir.
3. Kod denetler: alıntılar yorumda var mı; iyileştirme başlığının dayandığı yorumlar var ve o birime mi ait; en az 2
   yoruma ya da kodun "düşük" bulduğu bir soruya dayanıyor mu. Sorun varsa modele geri gönderilir (en fazla --tur kez).

Kullanım:
    python agent.py                                     # örnek anketlerle dener
    python agent.py --girdi anketler.xlsx --hedef 4 --min-n 15
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
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
TEMALAR = ["Hekim iletişimi / ilgisi", "Hemşire / bakım", "Bekleme süresi", "Randevu", "Bilgilendirme", "Temizlik / hijyen", "Yemek", "Personel tutumu",
           "Fiziki ortam / konfor", "Otopark / ulaşım", "Ücret / faturalama", "Diğer"]
META = {"id": ("Anket No", "No", "Form No"), "tarih": ("Tarih", "Anket Tarihi"), "birim": ("Birim", "Bölüm", "Klinik"), "hekim": ("Hekim", "Doktor"),
        "yorum": ("Yorum", "Görüş", "Açıklama", "Öneri")}
UNVAN = {"uzm", "dr", "op", "prof", "doc", "yrd", "yard", "dt", "uz"}
SEMA = {
    "type": "object",
    "properties": {
        "yorumlar": {"type": "array", "items": {"type": "object", "properties": {
            "id": S, "temalar": {"type": "array", "items": {"type": "string", "enum": TEMALAR}}, "duygu": {"type": "string", "enum": ["Olumlu", "Nötr", "Olumsuz"]}, "alinti": S},
            "required": ["id", "temalar", "duygu", "alinti"], "additionalProperties": False}},
        "iyilestirmeler": {"type": "array", "items": {"type": "object", "properties": {
            "birim": S, "baslik": S, "gerekce": S, "dayanak": {"type": "array", "items": S}, "dusuk_soru": S, "oneri": S},
            "required": ["birim", "baslik", "gerekce", "dayanak", "dusuk_soru", "oneri"], "additionalProperties": False}},
    },
    "required": ["yorumlar", "iyilestirmeler"],
    "additionalProperties": False,
}


def anketleri_oku(yol: Path) -> tuple[list[dict], list[str]]:
    s = tablo_oku(yol)
    if not s:
        raise llm.LLMHatasi(f"{yol.name}: tablo boş.")
    bas = [str(x or "").strip() for x in s[0]]
    meta_idx = {a: next((i for i, b in enumerate(bas) if katla(b) in {katla(e) for e in es}), None) for a, es in META.items()}
    if meta_idx["birim"] is None:
        raise llm.LLMHatasi(f"{yol.name}: 'Birim' sütunu gerekli. Başlıklar: {bas}")
    sorular = []
    for i, b in enumerate(bas):
        if i in meta_idx.values() or not b:
            continue
        degerler = [r[i] for r in s[1:] if i < len(r) and r[i] not in (None, "")]
        try:
            if degerler and all(1 <= float(sayi_norm(v)) <= 5 for v in degerler):
                sorular.append((i, b))
        except ValueError:
            pass
    if not sorular:
        raise llm.LLMHatasi(f"{yol.name}: 1–5 puanlı soru sütunu bulunamadı.")
    out = []
    for j, r in enumerate(s[1:], 1):
        al = lambda a: r[meta_idx[a]] if meta_idx[a] is not None and meta_idx[a] < len(r) and r[meta_idx[a]] is not None else ""  # noqa: E731
        p = {b: float(sayi_norm(r[i])) for i, b in sorular if i < len(r) and r[i] not in (None, "")}
        out.append({"id": str(al("id") or f"A{j}"), "tarih": tarih_coz(al("tarih")), "birim": str(al("birim")).strip(), "hekim": str(al("hekim")).strip(),
                    "puan": p, "yorum": str(al("yorum")).strip()})
    return out, [b for _, b in sorular]


def hekim_cikar(metin: str, hekimler: set[str]) -> str:
    out = metin
    parcalar = set()
    for h in hekimler:
        for p in re.split(r"[\s.]+", h):
            if len(p) >= 3 and katla(p) not in UNVAN:
                parcalar.add(p)
    for p in sorted(parcalar, key=len, reverse=True):
        out = re.sub(rf"\b{re.escape(p)}\b", "[HEKİM]", out, flags=re.I)
    return out


def ozet(anket: list[dict], sorular: list[str], anahtar: str) -> list[dict]:
    gr = defaultdict(list)
    for a in anket:
        if a[anahtar]:
            gr[(a["birim"], a[anahtar]) if anahtar == "hekim" else a[anahtar]].append(a)
    out = []
    for k, l in gr.items():
        ort = {}
        for q in sorular:
            v = [a["puan"][q] for a in l if q in a["puan"]]
            ort[q] = round(sum(v) / len(v), 2) if v else None
        tum = [p for a in l for p in a["puan"].values()]
        out.append({"anahtar": k, "n": len(l), "ort": ort, "genel": round(sum(tum) / len(tum), 2) if tum else None,
                    "memnun": round(100 * sum(1 for p in tum if p >= 4) / len(tum), 1) if tum else None})
    return sorted(out, key=lambda x: -(x["genel"] or 0))


def dusukler(birim_ozet: list[dict], hedef: float) -> dict[str, list[str]]:
    return {b["anahtar"]: [q for q, v in b["ort"].items() if v is not None and v < hedef] for b in birim_ozet}


def denetle(t: dict, yorumlu: list[dict], dusuk: dict[str, list[str]], hekimler: set[str]) -> list[str]:
    sorun, y = [], {a["id"]: a for a in yorumlu}
    sonuc = {x["id"]: x for x in t["yorumlar"]}
    for a in yorumlu:
        x = sonuc.get(a["id"])
        if not x:
            sorun.append(f"{a['id']}: yorum sınıflandırılmadı")
        elif not metinde_var(x["alinti"], hekim_cikar(a["yorum"], hekimler)):
            sorun.append(f"{a['id']}: alıntı yorumda yok: '{x['alinti'][:50]}'")
        elif not x["temalar"]:
            sorun.append(f"{a['id']}: tema yok")
    for i, x in enumerate(t["iyilestirmeler"], 1):
        ad = f"İyileştirme {i} ({x['birim']} · {x['baslik'][:40]})"
        if x["birim"] not in dusuk:
            sorun.append(f"{ad}: birim ankette yok")
            continue
        yok = [d for d in x["dayanak"] if d not in y]
        baska = [d for d in x["dayanak"] if d in y and y[d]["birim"] != x["birim"]]
        if yok:
            sorun.append(f"{ad}: dayanak {yok} yorumlu anketler arasında yok")
        if baska:
            sorun.append(f"{ad}: dayanak {baska} başka birime ait")
        gecerli = [d for d in x["dayanak"] if d in y and y[d]["birim"] == x["birim"]]
        if len(gecerli) < 2 and x["dusuk_soru"] not in dusuk[x["birim"]]:
            sorun.append(f"{ad}: tek yoruma dayanıyor ve '{x['dusuk_soru'] or '-'}' bu birimin düşük puanlı sorularından değil "
                         f"(düşük sorular: {', '.join(dusuk[x['birim']]) or 'yok'})")
    return sorun


def uret(yorumlu: list[dict], birim_ozet: list[dict], dusuk: dict, hedef: float, hekimler: set[str], duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    puanlar = "\n".join(f"{b['anahtar']} (n={b['n']}): " + ", ".join(f"{q} {v:.2f}" for q, v in b["ort"].items() if v is not None)
                        + f" · düşük (< {hedef:g}): {', '.join(dusuk[b['anahtar']]) or 'yok'}" for b in birim_ozet)
    yor = "\n".join(f'<yorum id="{a["id"]}" birim="{a["birim"]}">{llm.maskele(hekim_cikar(a["yorum"], hekimler))}</yorum>' for a in yorumlu)
    mesaj = f"<birim_puanlari>\n{puanlar}\n</birim_puanlari>\n<temalar>{'; '.join(TEMALAR)}</temalar>\n<yorumlar>\n{yor}\n</yorumlar>"
    return llm.json_iste(sistem, mesaj + duzeltme_blogu(duzeltme, onceki), SEMA)


def rapor_yaz(cikti, anket, sorular, birim_oz, hekim_oz, dusuk, t, yorumlu, hedef, min_n, hekimler, kalan, tur) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Birim Özeti"
    baslik(ws, ["Birim", "Yanıt", *sorular, "Genel Ort.", "Memnuniyet % (4–5)"], (26, 7, *[13] * len(sorular), 10, 16))
    for b in birim_oz:
        ws.append([b["anahtar"], b["n"], *[b["ort"][q] for q in sorular], b["genel"], b["memnun"]])
        for j, q in enumerate(sorular, 3):
            if q in dusuk[b["anahtar"]]:
                ws.cell(ws.max_row, j).fill = KIRMIZI
    ws.append([])
    ws.append([f"Kırmızı: ortalaması {hedef:g} hedefinin altında. Memnuniyet %: 4 ya da 5 verilen yanıtların tüm yanıtlara oranı."])
    h = wb.create_sheet("Hekim Özeti")
    baslik(h, ["Birim", "Hekim", "Yanıt", *sorular, "Genel Ort.", "Not"], (24, 22, 7, *[13] * len(sorular), 10, 30))
    for x in sorted(hekim_oz, key=lambda x: (x["n"] < min_n, -(x["genel"] or 0))):
        az = x["n"] < min_n
        h.append([*x["anahtar"], x["n"], *[None if az else x["ort"][q] for q in sorular], None if az else x["genel"],
                  f"Yetersiz örnek (n < {min_n}); değerlendirmeye alınmadı" if az else ""])
        if az:
            boya(h, [len(sorular) + 5], SARI)
    i = wb.create_sheet("İyileştirme Başlıkları", 0)
    baslik(i, ["Birim", "Başlık", "Gerekçe", "Dayanak Yorumlar", "Düşük Puanlı Soru", "Öneri", "Sorumlu", "Termin"], (24, 30, 46, 16, 18, 50, 12, 10))
    for x in t["iyilestirmeler"]:
        i.append([x["birim"], x["baslik"], x["gerekce"], ", ".join(x["dayanak"]), x["dusuk_soru"], x["oneri"], "", ""])
        boya(i, [2, 3, 6], AI_DOLGU)
        boya(i, [7, 8], SARI)
    son_not(i, "Önemli: Başlıklar yapay zekâ önerisidir; kalite birimi puanlar ve yorumlarla birlikte değerlendirmelidir.", kalan, tur)
    hizala(i)
    y = wb.create_sheet("Yorumlar")
    sonuc = {x["id"]: x for x in t["yorumlar"]}
    baslik(y, ["Anket No", "Birim", "Yorum", "Temalar", "Duygu", "Alıntı"], (9, 24, 60, 34, 9, 34))
    for a in yorumlu:
        x = sonuc.get(a["id"], {})
        y.append([a["id"], a["birim"], hekim_cikar(a["yorum"], hekimler), ", ".join(x.get("temalar", [])), x.get("duygu", ""), x.get("alinti", "")])
        boya(y, [4, 5], AI_DOLGU)
    hizala(y)
    tm = wb.create_sheet("Tema Dağılımı")
    baslik(tm, ["Birim", "Tema", "Olumlu", "Olumsuz", "Nötr"], (24, 26, 8, 8, 8))
    say = defaultdict(Counter)
    for a in yorumlu:
        for te in sonuc.get(a["id"], {}).get("temalar", []):
            say[(a["birim"], te)][sonuc[a["id"]]["duygu"]] += 1
    for (b, te), c in sorted(say.items(), key=lambda x: (x[0][0], -x[1]["Olumsuz"])):
        tm.append([b, te, c["Olumlu"], c["Olumsuz"], c["Nötr"]])
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(girdi: Path, cikti: Path, hedef: float = 4.0, min_n: int = 10, tur: int = 2, evet: bool = False) -> dict:
    anket, sorular = anketleri_oku(girdi)
    hekimler = {a["hekim"] for a in anket if a["hekim"]}
    birim_oz, hekim_oz = ozet(anket, sorular, "birim"), ozet(anket, sorular, "hekim")
    dusuk = dusukler(birim_oz, hedef)
    yorumlu = [a for a in anket if a["yorum"]]
    print(f"[OK] {len(anket)} anket · {len(sorular)} soru · {len(birim_oz)} birim · {len(yorumlu)} yorum")
    if yorumlu:
        llm.onay_al(f"{len(yorumlu)} yorum (birim adıyla; hekim adları [HEKİM], telefon / e-posta maskeli) ve birim puan ortalamaları gönderilecek.", evet)
        t, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(yorumlu, birim_oz, dusuk, hedef, hekimler, d, o),
                                             lambda x: denetle(x, yorumlu, dusuk, hekimler), tur)
    else:
        t, kalan, yapilan = {"yorumlar": [], "iyilestirmeler": []}, [], 0
    rapor_yaz(cikti, anket, sorular, birim_oz, hekim_oz, dusuk, t, yorumlu, hedef, min_n, hekimler, kalan, yapilan)
    return {"anket": anket, "sorular": sorular, "birim": birim_oz, "hekim": hekim_oz, "dusuk": dusuk, "taslak": t, "kalan": kalan, "tur": yapilan}


def main(argv: list[str] | None = None) -> int:
    ortam_hazirla()
    p = argparse.ArgumentParser(description="Hasta memnuniyet anketlerinden birim ve hekim bazında rapor ve iyileştirme başlıkları çıkarır.")
    p.add_argument("--girdi", type=Path, default=BURASI / "ornek_veri" / "anketler.csv",
                   help="Anketler (.xlsx/.csv): Birim ve 1–5 puanlı soru sütunları; Anket No, Tarih, Hekim, Yorum isteğe bağlı")
    p.add_argument("--hedef", type=float, default=4.0, help="Bu ortalamanın altındaki birim × soru düşük sayılır (varsayılan 4,0)")
    p.add_argument("--min-n", type=int, default=10, help="Hekim değerlendirmesi için en az yanıt (varsayılan 10)")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "memnuniyet_raporu.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    try:
        s = calistir(a.girdi, a.cikti, a.hedef, a.min_n, a.tur, a.evet)
    except llm.LLMHatasi as h:
        print(f"[X] {h}")
        return 1
    for b in s["birim"]:
        print(f"[OK] {b['anahtar']}: n={b['n']} · ort. {b['genel']} · memnuniyet %{b['memnun']}" + (f" · düşük: {', '.join(s['dusuk'][b['anahtar']])}" if s["dusuk"][b["anahtar"]] else ""))
    print(f"[OK] {len(s['taslak']['iyilestirmeler'])} iyileştirme başlığı")
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
