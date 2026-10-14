"""
Randevu Hatırlatma Mesajları — Workers / Workless AI Agent
Hasta Hizmetleri › Hasta Kayıt Kabul Görevlisi

Hasta verisi modele GÖNDERİLMEZ. Akış:
1. Kod, seçilen günün randevularını okur ve (bölüm, işlem, dil) kombinasyonlarını çıkarır; her kombinasyon için kurumun
   onaylı hazırlık talimatlarını (bölüm + işlem, bölüm geneli ve "Tümü" satırları) toplar.
2. Model yalnız kombinasyonları ve onaylı talimatları görür; her kombinasyon için yer tutuculu kısa bir hatırlatma
   şablonu yazar: {AD} {TARIH} {SAAT} {BOLUM} {HEKIM_IFADE} {KURUM} {ILETISIM}.
3. Kod şablonları denetler: zorunlu yer tutucular var mı, bilinmeyen yer tutucu var mı, talimattaki her sayı (saat, gün,
   dakika) şablonda aynen geçiyor mu, talimatta olmayan sayı ya da ilaç / doz gibi tıbbi yönlendirme eklenmiş mi, en uzun
   değerlerle doldurulduğunda --azami karakteri aşıyor mu. Sorun varsa modele geri gönderilir (en fazla --tur kez).
4. Kod şablonları hastaların bilgileriyle doldurur: Telefon;Mesaj CSV'si ve Excel (Şablonlar + 'Onay', Mesajlar, Kontrol).
   Mesaj gönderilmez; gönderim kurumun SMS sistemiyle, şablonlar onaylandıktan sonra yapılır.

Kullanım:
    python agent.py                                              # örnek randevularla (15.10.2026)
    python agent.py --randevular randevular.xlsx --talimatlar hazirlik.xlsx --tarih 16.10.2026 --kurum "Hastane Adı" --iletisim "0212 000 00 00"
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

import belge
import llm

BURASI = Path(__file__).resolve().parent
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
ZORUNLU = ("{AD}", "{TARIH}", "{SAAT}", "{BOLUM}", "{KURUM}", "{ILETISIM}")
IZINLI = set(ZORUNLU) | {"{HEKIM_IFADE}"}
TIBBI = {"ilac": ("ilac", "medic", "drug"), "doz": ("doz", "dose"), "insulin": ("insulin",), "aspirin": ("aspirin",), "kan sulandirici": ("kan sulandirici", "anticoag", "blood thinner"),
         "birak": ("ilacinizi birak", "birakin", "stop taking", "discontinue")}
DILLER = {"TR": "Türkçe", "EN": "İngilizce", "DE": "Almanca", "AR": "Arapça", "RU": "Rusça"}

SEMA = {
    "type": "object",
    "properties": {"sablonlar": {"type": "array", "items": {"type": "object", "properties": {"kimlik": {"type": "string"}, "metin": {"type": "string"}},
                                                           "required": ["kimlik", "metin"], "additionalProperties": False}}},
    "required": ["sablonlar"],
    "additionalProperties": False,
}


def katla(s) -> str:
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(s or "").translate(_TR).lower()).split())


def _tablo(yol: Path) -> list[dict]:
    if yol.suffix.lower() in (".xlsx", ".xlsm"):
        wb = load_workbook(yol, data_only=True, read_only=True)
        try:
            satir = [[("" if c is None else (c.strftime("%d.%m.%Y") if isinstance(c, (date, datetime)) else str(c))) for c in r] for r in wb.active.iter_rows(values_only=True)]
        finally:
            wb.close()
    else:
        metin = belge.txt_oku(yol)
        ayirac = max(";\t,", key=metin.splitlines()[0].count) if metin.strip() else ";"
        satir = list(csv.reader(metin.splitlines(), delimiter=ayirac))
    satir = [r for r in satir if any(str(c).strip() for c in r)]
    if not satir:
        return []
    b = [katla(c) for c in satir[0]]
    return [{b[j]: str(r[j]).strip() for j in range(min(len(b), len(r)))} | {"_satir": n} for n, r in enumerate(satir[1:], 2)]


def _al(r: dict, *adlar) -> str:
    return next((r[a] for a in adlar if r.get(a)), "")


def randevu_oku(yol: Path, gun: date) -> tuple[list[dict], list[str]]:
    l, notlar = [], []
    for r in _tablo(yol):
        t = _al(r, "tarih", "randevu tarihi")
        try:
            g = datetime.strptime(t[:10], "%d.%m.%Y").date()
        except ValueError:
            notlar.append(f"Satır {r['_satir']}: tarih '{t}' okunamadı; alınmadı.")
            continue
        if g != gun:
            continue
        saat = _al(r, "saat", "randevu saati")
        if not re.fullmatch(r"\d{1,2}[:.]\d{2}", saat) or not _al(r, "bolum", "poliklinik"):
            notlar.append(f"Satır {r['_satir']}: saat ya da bölüm okunamadı; alınmadı.")
            continue
        dil = (_al(r, "dil") or "TR").upper()
        l.append({"no": _al(r, "randevu no", "no"), "ad": _al(r, "hasta adi", "hasta", "ad soyad"), "telefon": _al(r, "telefon", "cep telefonu", "gsm"), "tarih": g,
                  "saat": saat.replace(".", ":"), "bolum": _al(r, "bolum", "poliklinik"), "hekim": _al(r, "hekim", "doktor"), "islem": _al(r, "islem", "randevu turu") or "Muayene",
                  "dil": dil if dil in DILLER else "TR"})
    return l, notlar


def talimat_oku(yol: Path | None) -> list[dict]:
    if yol is None:
        return []
    return [{"bolum": katla(_al(r, "bolum")), "islem": katla(_al(r, "islem")), "metin": _al(r, "hazirlik talimati", "talimat", "hazirlik")} for r in _tablo(yol) if _al(r, "hazirlik talimati", "talimat", "hazirlik")]


def kombinasyonlar(randevular: list[dict], talimatlar: list[dict]) -> list[dict]:
    out = {}
    for r in randevular:
        k = (katla(r["bolum"]), katla(r["islem"]), r["dil"])
        if k in out:
            r["sablon"] = out[k]["kimlik"]
            continue
        b, i = k[0], k[1]
        tl = [t["metin"] for t in talimatlar if (t["bolum"] in (b, "tumu") and t["islem"] in (i, "tumu", ""))
              and not (t["bolum"] == "tumu" and t["islem"] not in ("tumu", ""))]
        tl.sort(key=lambda m: next((n for n, t in enumerate(talimatlar) if t["metin"] == m), 0))
        out[k] = {"kimlik": f"S{len(out) + 1}", "bolum": r["bolum"], "islem": r["islem"], "dil": r["dil"], "talimatlar": tl}
        r["sablon"] = out[k]["kimlik"]
    return list(out.values())


def _sayilar(s: str) -> set[str]:
    return set(re.findall(r"\d+", re.sub(r"\{[A-Z_]+\}", "", s)))


def doldur(sablon: str, r: dict, kurum: str, iletisim: str) -> str:
    m = sablon
    deg = {"{AD}": r["ad"] or ("Değerli hastamız" if r["dil"] == "TR" else "Dear patient"), "{TARIH}": f"{r['tarih']:%d.%m.%Y}", "{SAAT}": r["saat"], "{BOLUM}": r["bolum"],
           "{HEKIM_IFADE}": f" ({r['hekim']})" if r["hekim"] else "", "{KURUM}": kurum, "{ILETISIM}": iletisim}
    for k, v in deg.items():
        m = m.replace(k, v)
    return " ".join(m.split())


def denetle(t: dict, kombler: list[dict], randevular: list[dict], kurum: str, iletisim: str, azami: int) -> list[str]:
    sorunlar, sab = [], {s["kimlik"]: s["metin"] for s in t["sablonlar"]}
    for k in kombler:
        metin = sab.get(k["kimlik"])
        ad = f"{k['kimlik']} ({k['bolum']} · {k['islem']} · {k['dil']})"
        if metin is None:
            sorunlar.append(f"{ad}: şablon yazılmamış")
            continue
        eksik = [y for y in ZORUNLU if y not in metin]
        if eksik:
            sorunlar.append(f"{ad}: zorunlu yer tutucu eksik {eksik}")
        yabanci = set(re.findall(r"\{[^}]*\}", metin)) - IZINLI
        if yabanci:
            sorunlar.append(f"{ad}: bilinmeyen yer tutucu {sorted(yabanci)}")
        kaynak = " ".join(k["talimatlar"])
        eksik_sayi, fazla_sayi = _sayilar(kaynak) - _sayilar(metin), _sayilar(metin) - _sayilar(kaynak)
        if eksik_sayi:
            sorunlar.append(f"{ad}: talimattaki sayı(lar) şablonda yok {sorted(eksik_sayi, key=int)} — süreler aynen korunmalı")
        if fazla_sayi:
            sorunlar.append(f"{ad}: talimatta olmayan sayı(lar) eklenmiş {sorted(fazla_sayi, key=int)}")
        km, kk = katla(metin), katla(kaynak)
        for kavram, kelimeler in TIBBI.items():
            if any(w in km for w in kelimeler) and not any(w in kk for w in TIBBI[kavram]):
                sorunlar.append(f"{ad}: onaylı talimatta olmayan tıbbi yönlendirme ('{kavram}') eklenmiş")
        en_uzun = max((len(doldur(metin, r, kurum, iletisim)) for r in randevular if r.get("sablon") == k["kimlik"]), default=0)
        if en_uzun > azami:
            sorunlar.append(f"{ad}: doldurulmuş mesaj {en_uzun} karakter, sınır {azami}; kısalt")
    for s in t["sablonlar"]:
        if s["kimlik"] not in {k["kimlik"] for k in kombler}:
            sorunlar.append(f"{s['kimlik']}: tanımsız şablon kimliği")
    return sorunlar


def uret(kombler: list[dict], azami: int, duzeltme: list[str] | None = None, onceki: dict | None = None) -> dict:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    liste = "\n".join(f"{k['kimlik']} | bölüm: {k['bolum']} | işlem: {k['islem']} | dil: {DILLER[k['dil']]} | onaylı talimatlar: "
                      + (" / ".join(k["talimatlar"]) or "yok") for k in kombler)
    parcalar = [f"<kombinasyonlar>\n{liste}\n</kombinasyonlar>", f"<azami_karakter>{azami}</azami_karakter>"]
    if duzeltme:
        parcalar.append("<duzeltme>\nÖnceki taslak: " + json.dumps(onceki, ensure_ascii=False) + "\nKodun bulduğu sorunlar:\n" + "\n".join(f"- {s}" for s in duzeltme) + "\n</duzeltme>")
    return llm.json_iste(sistem, "\n".join(parcalar), SEMA)


BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
KIRMIZI = PatternFill("solid", fgColor="FDE2E1")
SARI = PatternFill("solid", fgColor="FFF4CE")
YESIL = PatternFill("solid", fgColor="E3F5E1")
UST = Alignment(vertical="top", wrap_text=True)


def _baslik(ws, b, g):
    ws.append(b)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
    for j, w in enumerate(g):
        ws.column_dimensions[chr(65 + j)].width = w


def rapor_yaz(cikti: Path, kombler: list[dict], t: dict, mesajlar: list[dict], atlanan: list[str], kalan: list[str], tur: int) -> Path:
    sab = {s["kimlik"]: s["metin"] for s in t["sablonlar"]}
    wb = Workbook()
    ws = wb.active
    ws.title = "Şablonlar"
    _baslik(ws, ["Şablon", "Bölüm", "İşlem", "Dil", "Onaylı Talimatlar", "Şablon Metni", "Onay"], (8, 18, 18, 6, 60, 70, 10))
    for k in kombler:
        ws.append([k["kimlik"], k["bolum"], k["islem"], k["dil"], "\n".join(k["talimatlar"]), sab.get(k["kimlik"], ""), ""])
        ws.cell(ws.max_row, 7).fill = SARI
        for h in ws[ws.max_row]:
            h.alignment = UST
    ws.append([])
    ws.append(["Son kontrol: " + ("; ".join(kalan) if kalan else "Şablonlar yer tutucu, sayı, tıbbi ifade ve uzunluk kontrolünden geçti")])
    ws.cell(ws.max_row, 1).fill = KIRMIZI if kalan else YESIL
    ws.append([f"Düzeltme turu: {tur} · Model: {llm.kullanim_ozeti()} · {datetime.now():%d.%m.%Y %H:%M}"])
    ws.append(["Önemli: Şablonlar sorumlu hekim / kalite birimi onayından sonra kullanılmalıdır. Hasta verisi modele gönderilmedi."])
    m = wb.create_sheet("Mesajlar")
    _baslik(m, ["Randevu No", "Telefon", "Saat", "Bölüm", "Şablon", "Mesaj", "Karakter"], (11, 16, 7, 18, 8, 100, 9))
    for x in mesajlar:
        m.append([x["no"], x["telefon"], x["saat"], x["bolum"], x["sablon"], x["mesaj"], len(x["mesaj"])])
        m.cell(m.max_row, 6).alignment = UST
    if atlanan:
        k = wb.create_sheet("Gönderilemeyen")
        _baslik(k, ["Açıklama"], (100,))
        for a in atlanan:
            k.append([a])
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)
    yol = cikti.with_name(cikti.stem + "_sms.csv")
    with open(yol, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Telefon", "Mesaj"])
        for x in mesajlar:
            w.writerow([x["telefon"], x["mesaj"]])
    return yol


def calistir(randevu_yol: Path, talimat_yol: Path | None, cikti: Path, gun: date, kurum: str, iletisim: str, azami: int = 480, tur: int = 2, evet: bool = False) -> dict:
    randevular, notlar = randevu_oku(randevu_yol, gun)
    if not randevular:
        raise belge.BelgeHatasi(f"{gun:%d.%m.%Y} için okunabilen randevu yok.")
    kombler = kombinasyonlar(randevular, talimat_oku(talimat_yol))
    print(f"[OK] {gun:%d.%m.%Y}: {len(randevular)} randevu · {len(kombler)} şablon kombinasyonu (bölüm · işlem · dil)")
    llm.onay_al("Yalnız bölüm / işlem / dil kombinasyonları ve onaylı hazırlık talimatları gönderilecek; hasta adı, telefonu ve randevu saati gönderilmez.", evet)
    taslak = uret(kombler, azami)
    kalan = denetle(taslak, kombler, randevular, kurum, iletisim, azami)
    yapilan = 0
    while kalan and yapilan < tur:
        yapilan += 1
        print(f"[i] Düzeltme turu {yapilan}: {len(kalan)} sorun modele geri gönderiliyor")
        taslak = uret(kombler, azami, kalan, taslak)
        kalan = denetle(taslak, kombler, randevular, kurum, iletisim, azami)
    sab = {s["kimlik"]: s["metin"] for s in taslak["sablonlar"]}
    hatali = {k.split(" ")[0].rstrip(":") for k in kalan}
    mesajlar, atlanan = [], list(notlar)
    for r in sorted(randevular, key=lambda r: r["saat"]):
        if not r["telefon"]:
            atlanan.append(f"{r['no']} {r['saat']} {r['bolum']}: telefon yok; hasta aranmalı.")
            continue
        if r["sablon"] in hatali or r["sablon"] not in sab:
            atlanan.append(f"{r['no']} {r['saat']} {r['bolum']}: şablon ({r['sablon']}) kontrolden geçmedi; mesaj üretilmedi.")
            continue
        mesajlar.append({**r, "mesaj": doldur(sab[r["sablon"]], r, kurum, iletisim)})
    sms = rapor_yaz(cikti, kombler, taslak, mesajlar, atlanan, kalan, yapilan)
    return {"randevular": randevular, "kombler": kombler, "taslak": taslak, "kalan": kalan, "tur": yapilan, "mesajlar": mesajlar, "atlanan": atlanan, "sms": sms}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    ornek = BURASI / "ornek_veri"
    p = argparse.ArgumentParser(description="Randevulardan bölüm, hekim ve hazırlık bilgisi içeren hatırlatma mesajları hazırlar (hasta verisi modele gönderilmez).")
    p.add_argument("--randevular", type=Path, default=ornek / "randevular.csv", help="Randevu No, Hasta Adı, Telefon, Tarih, Saat, Bölüm, Hekim, İşlem, Dil (TR / EN / DE / AR / RU)")
    p.add_argument("--talimatlar", type=Path, help="Onaylı hazırlık talimatları: Bölüm, İşlem (boş / Tümü), Hazırlık Talimatı")
    p.add_argument("--tarih", help="Mesaj gönderilecek randevu günü GG.AA.YYYY (varsayılan yarın; örnek veride 15.10.2026)")
    p.add_argument("--kurum", default="", help="Kurum adı (mesajda geçer)")
    p.add_argument("--iletisim", default="", help="İptal / değişiklik için telefon")
    p.add_argument("--azami", type=int, default=480, help="Doldurulmuş mesajın en fazla karakter sayısı (varsayılan 480)")
    p.add_argument("--tur", type=int, default=2, help="Sorun kalırsa en fazla düzeltme turu (varsayılan 2)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "randevu_hatirlatma.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    ornekte = a.randevular == ornek / "randevular.csv"
    talimat = a.talimatlar or (ornek / "hazirlik_talimatlari.csv" if ornekte else None)
    try:
        gun = datetime.strptime(a.tarih, "%d.%m.%Y").date() if a.tarih else (date(2026, 10, 15) if ornekte else date.today() + timedelta(days=1))
    except ValueError:
        print("[X] --tarih GG.AA.YYYY olmalı")
        return 2
    kurum = a.kurum or ("Örnek Hastanesi" if ornekte else "")
    iletisim = a.iletisim or ("0212 000 00 00" if ornekte else "")
    if not kurum or not iletisim or a.azami < 100:
        print("[X] --kurum ve --iletisim zorunlu (örnek veri dışında); --azami en az 100 olmalı")
        return 2
    for y in (a.randevular, talimat):
        if y and not y.exists():
            print(f"[X] Dosya bulunamadı: {y}")
            return 1
    try:
        s = calistir(a.randevular, talimat, a.cikti, gun, kurum, iletisim, a.azami, a.tur, a.evet)
    except (llm.LLMHatasi, belge.BelgeHatasi) as h:
        print(f"[X] {h}")
        return 1
    print(f"[OK] {len(s['mesajlar'])} mesaj hazırlandı · gönderilemeyen {len(s['atlanan'])}")
    for k in s["kalan"]:
        print(f"[!] Kalan: {k}")
    for x in s["atlanan"]:
        print(f"[!] {x}")
    print(f"[OK] Rapor: {a.cikti.resolve()} (+ {s['sms'].name})")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
