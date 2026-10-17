"""
Mağaza Olay Tutanağı Taslağı — Gerisi Bende AI Agent
Perakende Mağaza › Kasa Sorumlusu (Mağaza Müdür Yardımcısı ve Kayıp Önleme Uzmanı da kullanır)

Kasa farkı, hırsızlık şüphesi, hasarlı / kayıp ürün gibi mağaza olaylarının kısa notlarından nesnel, imzaya hazır olay
tutanağı taslağı hazırlar.
1. Kod not dosyasını olaylara ayırır ("Olay: GG.AA.YYYY SS:DD | Yer" satırı), satırları numaralar. --gizle ile verilen kişi
   adları [KİŞİ-n] yapılır ve sağlayıcıya gönderilmez; telefon, e-posta, kimlik numarası maskelenir.
2. Model olay türünü, nesnel olay özetini, zaman çizelgesini, etkilenen ürün / tutarları, kasa sayım bilgisini, kanıtları,
   yapılan bildirimleri, yapılacakları ve eksik bilgileri yazar; her olay dayandığı not satırlarını gösterir.
3. Kod taslağı denetler: saat ve sayılar notta geçiyor mu; suçlayıcı ifade (hırsız, çaldı, suçlu, dolandırıcı, kasıtlı,
   yalan…) var mı; kişi adı yerine takma ad kullanılmış mı. Sorun varsa yalnız o olay modele geri gönderilir.
4. Kod kasa farkını kendisi hesaplar (sayılan − beklenen) ve notta yazan farkla karşılaştırır; ürün tutarı = adet × birim
   fiyat kontrolü yapar; tutanağı sabit şablonla (imza alanlarıyla) oluşturur.

Kullanım:
    python agent.py                                                    # örnek 3 olay
    python agent.py --girdi olaylar.txt --gizle "Ayşe Deneme" "Can Kurgu" --magaza "Örnek Market Kadıköy"
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
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
BASLIK = re.compile(r"^\s*Olay\s*:\s*(\d{1,2}\.\d{1,2}\.\d{4})(?:\s+(\d{1,2}:\d{2}))?\s*\|\s*(.+?)\s*$", re.I)
SUCLAMA = re.compile(r"\b(hırsız(?!lık)\w*|çal(dı|mış|an|ıyor|arak)\w*|suçlu\w*|dolandırıcı\w*|kasıtlı\w*|kasten|bilerek|yalan\w*|sahtekâr\w*|sahtekar\w*|"
                     r"zimmet\w*|aşır(dı|mış)\w*|yürüttü\w*|hileli)\b", re.I)
SES = re.compile(r"\b(alarm\w*|zil\w*|telefon\w*|siren\w*|korna\w*|müzik\w*|anons\w*)\s+$", re.I)    # "alarm çaldı" suçlama değildir


def suclama(metin: str) -> str | None:
    """Suçlayıcı ifade; 'çal-' fiili alarm / zil / telefon gibi bir sesten sonra geliyorsa sayılmaz."""
    for m in SUCLAMA.finditer(metin):
        if m.group(0).lower().startswith("çal") and SES.search(metin[:m.start()]):
            continue
        return m.group(0)
    return None
TUR = ["Kasa farkı", "Hırsızlık şüphesi", "Hasarlı ürün", "Kayıp ürün", "İade / indirim usulsüzlüğü şüphesi", "Müşteri / personel olayı", "Diğer"]


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9%]+", " ", str(s or "").replace("İ", "i").replace("I", "ı").translate(_TR).lower()).split())


def para(x) -> Decimal | None:
    if x in (None, ""):
        return None
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
    temiz = re.sub(r"\b\d{1,2}:\d{2}\b", " ", str(metin))
    return {abs(v) for x in re.findall(r"\d+(?:[.,]\d+)*", temiz) if (v := para(x)) is not None}


def saatler(metin: str) -> set[str]:
    return {f"{int(h):02d}:{m}" for h, m in re.findall(r"\b(\d{1,2}):(\d{2})\b", str(metin))}


def tl(x: Decimal) -> str:
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ----------------------------------------------------------------------------
# Girdi
# ----------------------------------------------------------------------------

def olaylari_oku(yol: Path, kisiler: dict[str, str]) -> list[dict]:
    if not yol.exists():
        raise llm.LLMHatasi(f"{yol}: not dosyası bulunamadı.")
    out = []
    for s in belge.metin_oku(yol).splitlines():
        m = BASLIK.match(s)
        if m:
            try:
                t = datetime.strptime(m[1] + (" " + m[2] if m[2] else ""), "%d.%m.%Y %H:%M" if m[2] else "%d.%m.%Y")
            except ValueError as h:
                raise llm.LLMHatasi(f"'{s.strip()}': tarih okunamadı (GG.AA.YYYY SS:DD).") from h
            out.append({"no": f"O{len(out) + 1}", "zaman": t, "saat_var": bool(m[2]), "yer": m[3], "satirlar": []})
        elif s.strip() and out:
            out[-1]["satirlar"].append(maskele(s.strip(), kisiler))
    out = [o for o in out if o["satirlar"]]
    if not out:
        raise llm.LLMHatasi(f"{yol.name}: 'Olay: GG.AA.YYYY SS:DD | Yer' başlıklı olay yok.")
    return out


def takma_adlar(adlar: list[str]) -> dict[str, str]:
    return {a: f"[KİŞİ-{i}]" for i, a in enumerate(dict.fromkeys(x.strip() for x in adlar if x.strip()), 1)}


def maskele(metin: str, kisiler: dict[str, str]) -> str:
    for gercek, takma in sorted(kisiler.items(), key=lambda x: -len(x[0])):
        metin = re.sub(re.escape(gercek), takma, metin, flags=re.I)
    return llm.maskele(metin)


# ----------------------------------------------------------------------------
# Model
# ----------------------------------------------------------------------------

ZAMAN = {"type": "object", "properties": {"saat": S, "olay": S}, "required": ["saat", "olay"], "additionalProperties": False}
URUN = {"type": "object", "properties": {"urun": S, "adet": S, "birim_fiyat": S, "tutar": S}, "required": ["urun", "adet", "birim_fiyat", "tutar"], "additionalProperties": False}
SEMA = {"type": "object", "properties": {
    "baslik": S, "tur": {"type": "string", "enum": TUR}, "ozet": S, "zaman_cizelgesi": {"type": "array", "items": ZAMAN},
    "urunler": {"type": "array", "items": URUN},
    "kasa": {"type": "object", "properties": {"beklenen": S, "sayilan": S, "nottaki_fark": S}, "required": ["beklenen", "sayilan", "nottaki_fark"], "additionalProperties": False},
    "kanitlar": {"type": "array", "items": S}, "bildirimler": {"type": "array", "items": S}, "yapilacaklar": {"type": "array", "items": S},
    "eksik_bilgiler": {"type": "array", "items": S}, "satirlar": {"type": "array", "items": {"type": "integer"}}},
    "required": ["baslik", "tur", "ozet", "zaman_cizelgesi", "urunler", "kasa", "kanitlar", "bildirimler", "yapilacaklar", "eksik_bilgiler", "satirlar"],
    "additionalProperties": False}


def metinler(c: dict) -> list[tuple[str, str]]:
    out = [("Başlık", c["baslik"]), ("Özet", c["ozet"])] + [(f"Zaman {i}", z["olay"]) for i, z in enumerate(c["zaman_cizelgesi"], 1)]
    for ad, alan in (("Kanıt", "kanitlar"), ("Bildirim", "bildirimler"), ("Yapılacak", "yapilacaklar")):
        out += [(f"{ad} {i}", x) for i, x in enumerate(c[alan], 1)]
    return out


def denetle(o: dict, c: dict) -> list[str]:
    on, notu = o["no"], "\n".join([o["yer"], *o["satirlar"]])
    izinli = sayilar(notu) | {Decimal(o["zaman"].day), Decimal(o["zaman"].month), Decimal(o["zaman"].year)}
    izinli_saat = saatler(notu) | ({f"{o['zaman']:%H:%M}"} if o["saat_var"] else set())
    sorun = []
    if not c["satirlar"] or any(not 1 <= i <= len(o["satirlar"]) for i in c["satirlar"]):
        sorun.append(f"{on}: dayanak satırlar boş ya da geçersiz (1-{len(o['satirlar'])})")
    for ad, m in metinler(c):
        fazla = sorted(sayilar(m) - izinli, key=float)
        if fazla:
            sorun.append(f"{on}: {ad}: notta olmayan sayı: {', '.join(f'{x:g}' for x in fazla)}")
        if s := suclama(m):
            sorun.append(f"{on}: {ad}: suçlayıcı / yargılayıcı ifade ('{s}'); 'tespit edildi', 'şüphesi', 'görüldü' gibi nesnel dil kullanın")
    for i, z in enumerate(c["zaman_cizelgesi"], 1):
        if z["saat"].strip() and saatler(z["saat"]) - izinli_saat:
            sorun.append(f"{on}: Zaman {i}: saat '{z['saat']}' notta yazmıyor")
    for i, u in enumerate(c["urunler"], 1):
        for a in ("adet", "birim_fiyat", "tutar"):
            if u[a].strip() and (para(u[a]) is None or abs(para(u[a])) not in izinli):
                sorun.append(f"{on}: Ürün {i}: {a} '{u[a]}' notta yazmıyor (hesaplamayın; kod hesaplar)")
    for a in ("beklenen", "sayilan", "nottaki_fark"):
        v = c["kasa"][a]
        if v.strip() and (para(v) is None or abs(para(v)) not in izinli):
            sorun.append(f"{on}: Kasa: {a} '{v}' notta yazmıyor (farkı hesaplamayın; kod hesaplar)")
    if c["tur"] == "Kasa farkı" and not (c["kasa"]["beklenen"].strip() and c["kasa"]["sayilan"].strip()) and not c["eksik_bilgiler"]:
        sorun.append(f"{on}: kasa farkı olayında beklenen ve sayılan tutar yok; notta yoksa eksik bilgilere yazın")
    return sorun


def hesap(c: dict) -> list[tuple[str, str]]:
    """Kod: kasa farkı ve ürün tutarı kontrolleri."""
    b = []
    k = c["kasa"]
    be, sa, nf = para(k["beklenen"]), para(k["sayilan"]), para(k["nottaki_fark"])
    if be is not None and sa is not None:
        fark = sa - be
        b.append(("Bilgi", f"Kasa farkı (kodla): {tl(sa)} − {tl(be)} = {tl(fark)} TL ({'eksik' if fark < 0 else 'fazla' if fark > 0 else 'fark yok'})"))
        if nf is not None and abs(nf) != abs(fark):
            b.append(("Hata", f"Notta fark {tl(abs(nf))} TL yazıyor, hesaplanan {tl(abs(fark))} TL: sayımı tekrarlayın"))
    toplam = Decimal(0)
    for u in c["urunler"]:
        a, f, t = para(u["adet"]), para(u["birim_fiyat"]), para(u["tutar"])
        if a is not None and f is not None:
            hes = a * f
            toplam += hes
            if t is not None and t != hes:
                b.append(("Hata", f"{u['urun']}: {a:g} × {tl(f)} = {tl(hes)} TL, notta {tl(t)} TL"))
        elif t is not None:
            toplam += t
    if toplam:
        b.append(("Bilgi", f"Etkilenen ürünlerin tutarı (kodla): {tl(toplam)} TL"))
    return b


def uret(olaylar: list[dict], duzeltme=None, onceki=None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    sonuc = dict(onceki or {})
    hedef = olaylar if duzeltme is None else [o for o in olaylar if any(s.startswith(o["no"] + ":") for s in duzeltme)]
    for o in hedef:
        satir = "\n".join(f"{i}. {s}" for i, s in enumerate(o["satirlar"], 1))
        zaman = f"{o['zaman']:%d.%m.%Y %H:%M}" if o["saat_var"] else f"{o['zaman']:%d.%m.%Y}"
        mesaj = f"<olay tarih=\"{zaman}\" yer=\"{o['yer']}\">\n{satir}\n</olay>"
        d = [s.split(": ", 1)[1] for s in (duzeltme or []) if s.startswith(o["no"] + ":")]
        if d:
            mesaj += ("\n<duzeltme>\nÖnceki taslak: " + json.dumps(sonuc.get(o["no"]), ensure_ascii=False) + "\nKodun bulduğu sorunlar:\n"
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


def tutanak(o: dict, c: dict, bulgu: list, no: str, magaza: str) -> str:
    z = f"{o['zaman']:%d.%m.%Y %H:%M}" if o["saat_var"] else f"{o['zaman']:%d.%m.%Y}"
    l = ["OLAY TUTANAĞI", f"Tutanak No: {no}", f"Mağaza: {magaza or '..........'} · Yer: {o['yer']}", f"Olay tarihi / saati: {z}",
         f"Olay türü: {c['tur']}", f"Konu: {c['baslik']}", ""]
    bolumler = [("Olayın özeti", [c["ozet"]]),
                ("Zaman çizelgesi", [f"   {x['saat'] or '--:--'}  {x['olay']}" for x in c["zaman_cizelgesi"]]),
                ("Etkilenen ürünler", [f"   {u['urun']} · adet: {u['adet'] or '-'} · birim fiyat: {u['birim_fiyat'] or '-'} · tutar: {u['tutar'] or '-'}"
                                       for u in c["urunler"]]),
                ("Hesap ve kontroller (kod tarafından)", [f"   {m}" for _, m in bulgu])]
    bolumler += [(b, [f"   • {x}" for x in c[a]]) for b, a in (("Kanıtlar", "kanitlar"), ("Yapılan bildirimler", "bildirimler"),
                                                               ("Yapılacaklar", "yapilacaklar"), ("Tamamlanması gereken bilgiler", "eksik_bilgiler"))]
    for i, (b, satirlar) in enumerate([x for x in bolumler if x[1]], 1):          # boş bölümler atlanır, numara sıralı kalır
        l += [f"{i}. {b}", *satirlar, ""]
    l += ["Bu tutanak olay yerinde gözlenen ve kayıtlara geçen bilgilere dayanır; hiçbir kişi hakkında suç isnadı içermez.",
          "Tutanak .... nüsha olarak düzenlenmiş ve aşağıda imzası bulunanlarca okunarak imzalanmıştır.", "",
          "Düzenleyen (ad soyad, görev, imza): ....................................",
          "Tanık (ad soyad, görev, imza): ....................................",
          "Mağaza yöneticisi (ad soyad, imza): ...................................."]
    return "\n".join(l)


def _baslik(ws, basliklar, genislik) -> None:
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font, h.alignment = BASLIK_DOLGU, BASLIK_YAZI, UST
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def rapor_yaz(cikti: Path, olaylar: list[dict], sonuc: dict, kalan: list[str], tur: int, magaza: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Olaylar"
    _baslik(ws, ["Tutanak No", "Tarih", "Yer", "Olay Türü", "Konu", "Etkilenen Tutar", "Kasa Farkı", "Hesap Bulguları", "Tutanak", "Son Kontrol"],
            (16, 16, 30, 18, 40, 12, 12, 50, 90, 30))
    metinler_ = []
    for i, o in enumerate(olaylar, 1):
        c = sonuc[o["no"]]
        no = f"OT-{o['zaman']:%y%m%d}-{i:02d}"
        b = o["_bulgu"]
        t = tutanak(o, c, b, no, magaza)
        metinler_.append(t)
        kasa = next((m for _, m in b if m.startswith("Kasa farkı (kodla)")), "")
        tutar = next((m for _, m in b if m.startswith("Etkilenen ürünlerin tutarı")), "")
        sor = [x.split(": ", 1)[1] for x in kalan if x.startswith(o["no"] + ":")]
        ws.append([no, o["zaman"], o["yer"], c["tur"], c["baslik"], tutar.split(": ", 1)[1] if tutar else "", kasa.split(" = ")[1].split(" (")[0] if kasa else "",
                   "\n".join(f"{x}: {m}" for x, m in b), t, "; ".join(sor) or "Geçti"])
        n = ws.max_row
        ws.cell(n, 2).number_format = "DD.MM.YYYY HH:MM"
        for j in (4, 5, 9):
            ws.cell(n, j).fill = AI_DOLGU
        if any(x == "Hata" for x, _ in b):
            ws.cell(n, 8).fill = KIRMIZI
        ws.cell(n, 10).fill = KIRMIZI if sor else YESIL
        for h in ws[n]:
            h.alignment = UST
    k = wb.create_sheet("Kontrol")
    k.append([f"Düzeltme turu: {tur} · Model: {llm.kullanim_ozeti()} · {datetime.now():%d.%m.%Y %H:%M}"])
    k.append(["Son kontrol: " + ("; ".join(kalan) if kalan else "Kod denetiminden sorunsuz geçti")])
    k.cell(2, 1).fill = KIRMIZI if kalan else YESIL
    k.append(["Önemli: Tutanaklar taslaktır. İmzalamadan önce olaya tanık olanlarla birlikte okuyun ve düzeltin. Kişiler hakkında yargı "
              "içermemelidir; şüpheli olaylarda kamera kaydını ve kanıtları koruyun, adli süreç için hukuk / güvenlik biriminize danışın."])
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)
    txt = cikti.with_suffix(".txt")
    txt.write_text("\n\n" + ("\n\n" + "=" * 72 + "\n\n").join(metinler_), encoding="utf-8")
    return txt


def calistir(girdi: Path, cikti: Path, gizle: list[str] | None = None, magaza: str = "", tur: int = 2, evet: bool = False) -> dict:
    kisiler = takma_adlar(gizle or [])
    olaylar = olaylari_oku(girdi, kisiler)
    print(f"[OK] {len(olaylar)} olay · {sum(len(o['satirlar']) for o in olaylar)} not satırı · gizlenen ad: {len(kisiler)}")
    llm.onay_al(f"{len(olaylar)} olay notu gönderilecek{f'; {len(kisiler)} kişi adı takma adla' if kisiler else ''} (telefon, e-posta, TCKN maskeli). "
                "Not içinde --gizle ile verilmeyen adlar varsa onlar da gider.", evet)
    sonuc, kalan, yapilan = duzeltme_dongusu(lambda d, o: uret(olaylar, d, o), lambda x: [s for o in olaylar for s in denetle(o, x[o["no"]])], tur)
    for o in olaylar:
        o["_bulgu"] = hesap(sonuc[o["no"]])
    sonuc = geri_ac(sonuc, kisiler)
    txt = rapor_yaz(cikti, olaylar, sonuc, kalan, yapilan, magaza)
    return {"olaylar": olaylar, "sonuc": sonuc, "kalan": kalan, "tur": yapilan, "txt": txt}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    p = argparse.ArgumentParser(description="Mağaza olay notlarından nesnel, imzaya hazır olay tutanağı taslakları hazırlar.")
    p.add_argument("--girdi", type=Path, default=ORNEK / "olay_notlari.txt", help="Notlar (.txt/.docx); her olay 'Olay: GG.AA.YYYY SS:DD | Yer' ile başlar")
    p.add_argument("--gizle", nargs="*", default=None, metavar="AD", help="Sağlayıcıya gönderilmeyecek kişi adları (personel, müşteri)")
    p.add_argument("--magaza", default="", help="Tutanaktaki mağaza adı")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "olay_tutanaklari.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    ornek = a.girdi == ORNEK / "olay_notlari.txt"
    gizle = a.gizle if a.gizle is not None else (["Selin Deneme"] if ornek else [])
    try:
        s = calistir(a.girdi, a.cikti, gizle, a.magaza or ("Örnek Market Kadıköy" if ornek else ""), a.tur, a.evet)
    except (llm.LLMHatasi, belge.BelgeHatasi) as h:
        print(f"[X] {h}")
        return 1
    for o in s["olaylar"]:
        print(f"[OK] {o['no']} · {s['sonuc'][o['no']]['tur']} · {s['sonuc'][o['no']]['baslik'][:60]}")
        for x, m in o["_bulgu"]:
            if x == "Hata":
                print(f"[!]   {m}")
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    print(f"[OK] Rapor: {a.cikti.resolve()} (+ {s['txt'].name})")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
