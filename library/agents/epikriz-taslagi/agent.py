"""
Epikriz Taslağı — Gerisi Bende AI Agent
Tıbbi Sekreterlik › Tıbbi Sekreter

1. Kod hasta kaydını, hekim notlarını (tarihli notlar N1, N2 …; "Tanılar" ve "Çıkış İlaçları" bölümleri) ve tetkik
   sonuçlarını okur. Hasta adı, TC kimlik no ve protokol no modele gönderilmez.
2. Kod hesaplar: yatış süresi, her tetkik sonucunun referans aralığına göre durumu (yüksek / düşük), aynı tetkikin ilk ve
   son değeri.
3. Model epikrizin anlatı bölümlerini yazar: başvuru şikâyeti ve öykü, fizik muayene, yapılan işlemler, klinik seyir,
   çıkıştaki durum, öneriler. Her bölüm dayandığı not (N) ve tetkik (T) numaralarını gösterir.
4. Kod denetler:
   - Bölümlerdeki sayılar (değer, doz, tarih) notlarda ya da tetkik tablosunda var mı
   - Metindeki tanı kodları (ICD-10) "Tanılar" bölümünde var mı; tarihler yatış-çıkış aralığında mı (öneriler hariç)
   - Referans dışı her tetkik klinik seyirde ya da tetkik bölümünde anılmış mı
   - Notlarda olmayan kesinlik ifadesi ("tamamen iyileşti", "kesin tanı") var mı
   Sorun varsa modele geri gönderilir (en fazla --tur kez).
5. Kod epikrizi hasta kimliği, tanılar, tetkik tablosu ve çıkış ilaçlarıyla (notlardaki gibi, aynen) birleştirir.
   **Hekim onayı olmadan kullanılamaz.**

Kullanım:
    python agent.py                                                    # örnek (kurgusal) apendektomi yatışı
    python agent.py --hasta hasta.txt --notlar notlar.docx --tetkik tetkikler.xlsx
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
HASTA = {"ad": "Hasta Adı", "tc": "TC Kimlik No", "dogum": "Doğum Tarihi", "cinsiyet": "Cinsiyet", "protokol": "Protokol No", "servis": "Servis",
         "yatis": "Yatış Tarihi", "cikis": "Çıkış Tarihi", "hekim": "Sorumlu Hekim"}
TETKIK = {"tarih": ("Tarih",), "ad": ("Tetkik", "Test"), "sonuc": ("Sonuç", "Değer"), "birim": ("Birim",), "ref": ("Referans", "Referans Aralığı")}
ICD = re.compile(r"\b[A-TV-Z]\d{2}(?:\.\d{1,2})?\b")
KESIN = re.compile(r"tamamen iyiles\w*|tam sifa|kesin tani\w*|kesinlikle|hicbir (sorun|sikayet)\w*|sorunsuz iyiles\w*")
BOLUMLER = ["oyku", "fizik_muayene", "yapilan_islemler", "klinik_seyir", "cikis_durumu", "oneriler"]
BASLIKLAR = {"oyku": "Başvuru Şikâyeti ve Öykü", "fizik_muayene": "Fizik Muayene", "yapilan_islemler": "Yapılan İşlemler ve Tedavi",
             "klinik_seyir": "Klinik Seyir", "cikis_durumu": "Çıkıştaki Durum", "oneriler": "Öneriler"}
B = {"type": "object", "properties": {"metin": S, "dayanak": {"type": "array", "items": S}}, "required": ["metin", "dayanak"], "additionalProperties": False}

SEMA = {"type": "object", "properties": {k: B for k in BOLUMLER}, "required": BOLUMLER, "additionalProperties": False}


def hasta_oku(yol: Path) -> dict:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: hasta kaydı bulunamadı.")
    h = {a: "" for a in HASTA}
    for s in belge.metin_oku(yol).splitlines():
        a, _, d = s.partition(":")
        k = next((k for k, v in HASTA.items() if katla(v) == katla(a)), None)
        if k:
            h[k] = d.strip()
    h["yatis_v"], h["cikis_v"] = tarih_coz(h["yatis"]), tarih_coz(h["cikis"])
    if not h["yatis_v"] or not h["cikis_v"] or h["cikis_v"] < h["yatis_v"]:
        raise llm.LLMHatasi(f"{yol.name}: 'Yatış Tarihi' ve 'Çıkış Tarihi' (GG.AA.YYYY, çıkış yatıştan önce olamaz) gerekli.")
    return h


def notlari_oku(yol: Path) -> dict:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: hekim notları bulunamadı.")
    n = {"tanilar": [], "notlar": [], "ilaclar": []}
    bolum = None
    for s in belge.metin_oku(yol).splitlines():
        if not s.strip():
            continue
        k = katla(s.rstrip(":"))
        if k in {"tanilar", "notlar", "cikis ilaclari"}:
            bolum = {"tanilar": "tanilar", "notlar": "notlar", "cikis ilaclari": "ilaclar"}[k]
            continue
        if not bolum:
            continue
        x = s.strip().lstrip("-•* ").strip()
        if bolum == "notlar":
            m = re.match(r"(\d{2}\.\d{2}\.\d{4})\s*[:\-–]\s*(.+)", x)
            n["notlar"].append({"id": f"N{len(n['notlar']) + 1}", "tarih": m.group(1) if m else "", "metin": m.group(2) if m else x})
        else:
            n[bolum].append(x)
    if not n["notlar"]:
        raise llm.LLMHatasi(f"{yol.name}: 'Notlar' bölümü gerekli ('- GG.AA.YYYY: …' satırları).")
    return n


def ref_coz(ref: str) -> tuple[float | None, float | None]:
    r = str(ref or "").replace(",", ".")
    m = re.match(r"\s*([\d.]+)\s*[-–]\s*([\d.]+)", r)
    if m:
        return float(m.group(1)), float(m.group(2))
    m = re.match(r"\s*<\s*([\d.]+)", r)
    if m:
        return None, float(m.group(1))
    m = re.match(r"\s*>\s*([\d.]+)", r)
    return (float(m.group(1)), None) if m else (None, None)


def tetkik_oku(yol: Path) -> list[dict]:
    out = []
    for r in kayitlar(yol, TETKIK, ("tarih", "ad", "sonuc")):
        if not r["ad"]:
            continue
        sonuc = str(r["sonuc"]).strip()
        try:
            v = None if sonuc.startswith(("<", ">")) else float(sonuc.replace(",", "."))
        except ValueError:
            v = None
        alt, ust = ref_coz(r["ref"])
        durum = "" if v is None else ("Yüksek" if ust is not None and v > ust else ("Düşük" if alt is not None and v < alt else "Normal"))
        out.append({"id": f"T{len(out) + 1}", "tarih": tarih_coz(r["tarih"]), "ad": str(r["ad"]).strip(), "sonuc": sonuc, "deger": v,
                    "birim": str(r["birim"] or ""), "ref": str(r["ref"] or ""), "durum": durum})
    return out


def tetkik_adi(ad: str) -> str:
    """'Lökosit (WBC)' → 'lokosit'; metinde arama için ilk kelime."""
    return katla(re.sub(r"\(.*?\)", "", ad)).split()[0] if katla(re.sub(r"\(.*?\)", "", ad)) else katla(ad)


def gizle(metin: str, h: dict) -> str:
    for d in (h["ad"], h["protokol"]):
        if d:
            metin = re.sub(re.escape(d), "{HASTA}" if d == h["ad"] else "[PROTOKOL]", metin, flags=re.I)
    return llm.maskele(metin)


def denetle(t: dict, h: dict, n: dict, tet: list[dict]) -> list[str]:
    sorun = []
    nid = {x["id"]: x for x in n["notlar"]}
    tid = {x["id"]: x for x in tet}
    kaynak_tum = " ".join(f"{x['tarih']} {x['metin']}" for x in n["notlar"]) + " " + " ".join(n["ilaclar"]) + " " + " ".join(n["tanilar"])
    kaynak_tum += " " + " ".join(f"{x['tarih']:%d.%m.%Y} {x['sonuc']} {x['birim']} {x['ref']}" for x in tet if x["tarih"])
    izinli = sayilar(kaynak_tum + f" {h['yatis']} {h['cikis']} {(h['cikis_v'] - h['yatis_v']).days}")
    tanilar = " ".join(n["tanilar"])
    for k in BOLUMLER:
        b = t[k]
        ad = BASLIKLAR[k]
        gec = [d for d in b["dayanak"] if d not in nid and d not in tid]
        if not b["dayanak"] or gec:
            sorun.append(f"{ad}: dayanak yok ya da geçersiz {gec}")
        fazla = sorted(sayilar(ICD.sub(" ", b["metin"])) - izinli)
        if fazla:
            sorun.append(f"{ad}: notlarda / tetkiklerde olmayan sayı: {', '.join(fazla)}")
        for kod in ICD.findall(b["metin"]):
            if kod not in tanilar:
                sorun.append(f"{ad}: '{kod}' tanı kodu 'Tanılar' bölümünde yok")
        if k != "oneriler":
            for d in re.findall(r"\b\d{2}\.\d{2}\.\d{4}\b", b["metin"]):
                g = tarih_coz(d)
                if g and not (h["yatis_v"] <= g <= h["cikis_v"]):
                    sorun.append(f"{ad}: {d} yatış-çıkış ({h['yatis']}-{h['cikis']}) aralığında değil")
        km = katla(b["metin"])
        m = KESIN.search(km)
        if m and m.group(0) not in katla(kaynak_tum):
            sorun.append(f"{ad}: notlarda olmayan kesinlik ifadesi ('{m.group(0)}')")
    seyir = katla(t["klinik_seyir"]["metin"] + " " + t["yapilan_islemler"]["metin"] + " " + t["oyku"]["metin"])
    for x in tet:
        if x["durum"] in ("Yüksek", "Düşük") and tetkik_adi(x["ad"]) not in seyir:
            sorun.append(f"{x['id']}: referans dışı sonuç ({x['ad']} {x['sonuc']} {x['birim']}, {x['durum'].lower()}) metinde anılmamış")
    return list(dict.fromkeys(sorun))


def uret(h: dict, n: dict, tet: list[dict], duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    notlar = "\n".join(f"{x['id']} | {x['tarih']} | {gizle(x['metin'], h)}" for x in n["notlar"])
    tt = "\n".join(f"{x['id']} | {x['tarih']:%d.%m.%Y} | {x['ad']} | {x['sonuc']} {x['birim']} | ref {x['ref'] or '-'} | {x['durum'] or '-'}"
                   for x in tet if x["tarih"])
    mesaj = (f"<yatis>Servis: {h['servis'] or '-'} · {h['cinsiyet'] or '-'} · yatış {h['yatis']} · çıkış {h['cikis']} · "
             f"{(h['cikis_v'] - h['yatis_v']).days} gün</yatis>\n<tanilar>\n" + "\n".join(n["tanilar"]) + "\n</tanilar>\n"
             f"<notlar>\n{notlar}\n</notlar>\n<tetkikler>\n{tt or '(yok)'}\n</tetkikler>\n<cikis_ilaclari>\n" + "\n".join(n["ilaclar"]) +
             "\n</cikis_ilaclari>" + duzeltme_blogu(duzeltme, onceki))
    return llm.json_iste(sistem, mesaj, SEMA)


def epikriz_metni(t: dict, h: dict, n: dict, tet: list[dict]) -> str:
    s = ["EPİKRİZ (TASLAK — HEKİM ONAYI ZORUNLUDUR)", "", f"Hasta: {h['ad']} · TC: {h['tc']} · Doğum tarihi: {h['dogum']} · Cinsiyet: {h['cinsiyet']}",
         f"Protokol: {h['protokol']} · Servis: {h['servis']}", f"Yatış: {h['yatis']} · Çıkış: {h['cikis']} · Yatış süresi: {(h['cikis_v'] - h['yatis_v']).days} gün",
         f"Sorumlu hekim: {h['hekim']}", "", "TANILAR"] + [f"  {x}" for x in n["tanilar"]]
    for k in BOLUMLER[:4]:
        s += ["", BASLIKLAR[k].upper(), t[k]["metin"]]
    s += ["", "TETKİK SONUÇLARI"] + [f"  {x['tarih']:%d.%m.%Y}  {x['ad']}: {x['sonuc']} {x['birim']} (ref. {x['ref'] or '-'})" +
                                      (f" [{x['durum']}]" if x["durum"] in ("Yüksek", "Düşük") else "") for x in tet if x["tarih"]]
    s += ["", BASLIKLAR["cikis_durumu"].upper(), t["cikis_durumu"]["metin"], "", "ÇIKIŞ İLAÇLARI"] + [f"  - {x}" for x in n["ilaclar"] or ["Yok"]]
    s += ["", BASLIKLAR["oneriler"].upper(), t["oneriler"]["metin"], "", "Hekim onayı: ……………………  Tarih: ……………  İmza: ……………"]
    return "\n".join(s)


def rapor_yaz(cikti: Path, t, h, n, tet, kalan, tur, metin) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Epikriz"
    baslik(ws, ["Bölüm", "Metin", "Dayanak", "Dayanak Metni"], (24, 90, 14, 70))
    kay = {x["id"]: f"{x['tarih']} {x['metin']}" for x in n["notlar"]} | {x["id"]: f"{x['ad']} {x['sonuc']} {x['birim']}" for x in tet}
    for k in BOLUMLER:
        ws.append([BASLIKLAR[k], t[k]["metin"], ", ".join(t[k]["dayanak"]), "\n".join(f"{d}: {kay.get(d, '?')}" for d in t[k]["dayanak"])])
        boya(ws, (2,), AI_DOLGU)
    son_not(ws, "Önemli: Taslaktır. Epikriz hekim tarafından okunup düzeltilmeden ve onaylanmadan hastaya verilemez, sisteme kaydedilemez.", kalan, tur)
    hizala(ws)
    w = wb.create_sheet("Tetkikler")
    baslik(w, ["No", "Tarih", "Tetkik", "Sonuç", "Birim", "Referans", "Durum"], (6, 12, 24, 10, 12, 12, 10))
    for x in tet:
        w.append([x["id"], x["tarih"], x["ad"], x["sonuc"], x["birim"], x["ref"], x["durum"]])
        w.cell(w.max_row, 2).number_format = "DD.MM.YYYY"
        if x["durum"] in ("Yüksek", "Düşük"):
            boya(w, (7,), SARI)
    hizala(w)
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)
    cikti.with_suffix(".txt").write_text(metin, encoding="utf-8")


def calistir(hasta_yol: Path, not_yol: Path, tetkik_yol: Path, cikti: Path, tur: int = 2, evet: bool = False) -> dict:
    h = hasta_oku(hasta_yol)
    n = notlari_oku(not_yol)
    tet = tetkik_oku(tetkik_yol) if tetkik_yol and tetkik_yol.exists() else []
    print(f"[OK] {h['servis'] or 'Servis'} · {h['yatis']}-{h['cikis']} · {len(n['notlar'])} not · {len(tet)} tetkik "
          f"({sum(x['durum'] in ('Yüksek', 'Düşük') for x in tet)} referans dışı)")
    llm.onay_al("Hekim notları ve tetkik sonuçları gönderilecek (hasta adı, TC kimlik no, protokol no ve doğum tarihi gönderilmez). Sağlık verisi "
                "içerdiğinden kurumunuzun izni olmadan çalıştırmayın.", evet)
    t, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(h, n, tet, d, o), lambda x: denetle(x, h, n, tet), tur)
    metin = epikriz_metni(t, h, n, tet)
    rapor_yaz(cikti, t, h, n, tet, kalan, yapilan, metin)
    return {"hasta": h, "notlar": n, "tetkikler": tet, "taslak": t, "kalan": kalan, "tur": yapilan, "metin": metin}


def main(argv: list[str] | None = None) -> int:
    ortam_hazirla()
    p = argparse.ArgumentParser(description="Hekim notları ve tetkik sonuçlarından epikriz taslağı oluşturur (hekim onayı zorunludur).")
    v = BURASI / "ornek_veri"
    p.add_argument("--hasta", type=Path, default=v / "hasta.txt",
                   help="Hasta kaydı (.txt/.docx): Hasta Adı, TC Kimlik No, Doğum Tarihi, Cinsiyet, Protokol No, Servis, Yatış Tarihi, Çıkış Tarihi, Sorumlu Hekim")
    p.add_argument("--notlar", type=Path, default=v / "notlar.txt", help="Hekim notları (.txt/.docx): 'Tanılar', 'Notlar' (- GG.AA.YYYY: …), 'Çıkış İlaçları' bölümleri")
    p.add_argument("--tetkik", type=Path, default=v / "tetkikler.csv", help="Tetkikler (CSV/Excel): Tarih, Tetkik, Sonuç, Birim, Referans")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "epikriz_taslagi.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    try:
        s = calistir(a.hasta, a.notlar, a.tetkik, a.cikti, a.tur, a.evet)
    except (llm.LLMHatasi, belge.BelgeHatasi) as h:
        print(f"[X] {h}")
        return 1
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    print(f"[OK] Epikriz taslağı: {a.cikti.with_suffix('.txt').resolve()} (hekim onayı zorunlu)")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
