"""
Taşeron Hakediş Kontrolü — Workers / Workless kod bloğu
İnşaat › Teknik Ofis › Teknik Ofis Mühendisi

Taşeronun sunduğu hakedişi; sözleşme birim fiyat cetveli, saha (yeşil defter) metrajı ve önceki onaylı hakedişle
karşılaştırır:
  - Birim fiyat sözleşmeden farklı mı, poz sözleşmede var mı (yeni fiyat/ek iş), birim aynı mı?
  - Aritmetik: önceki + bu dönem = toplam, miktar × birim fiyat = tutar.
  - Taşeronun "önceki" miktarı, bizim önceki onaylı kümülatifimize eşit mi?
  - Toplam miktar saha metrajını (tolerans dahil) ve sözleşme miktarını aşıyor mu? Negatif bu dönem var mı?
  - Onaylanabilir tutar: kümülatif miktar en çok saha metrajı (ve istenirse sözleşme miktarı) kadar, bu dönem en çok
    talep edilen kadar, sözleşme birim fiyatıyla; bu dönem onaylanan = onaylanan kümülatif − önceki onaylı.
Vergi ve kesintiler (KDV, tevkifat, stopaj, teminat, avans) için "Hakediş Hesaplama" paketini kullanın.
İnternete bağlanmaz.

Kullanım:
    python main.py                                         # örnek verilerle dener
    python main.py --sozlesme birim_fiyatlar.xlsx --hakedis taseron_hakedis_5.xlsx --saha saha_metraj.xlsx --onceki onayli_4.xlsx
    python main.py ... --tolerans 1 --sozlesme-siniri
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
SIFIR = Decimal(0)
KURUS = Decimal("0.01")


def kucuk(s) -> str:
    return str(s or "").replace("İ", "i").replace("I", "ı").lower().strip()


def katla(s) -> str:
    s = kucuk(s)
    for a, b in zip("ıöüşğç", "iousgc"):
        s = s.replace(a, b)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def para(x) -> Decimal:
    if x in (None, ""):
        return SIFIR
    if isinstance(x, (int, float)):
        return Decimal(str(x))
    s = str(x).strip().replace("TL", "").replace("₺", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        return Decimal(s)
    except InvalidOperation:
        return SIFIR


def yuvarla(x: Decimal) -> Decimal:
    return x.quantize(KURUS, rounding=ROUND_HALF_UP)


def tl(x: Decimal) -> str:
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def poz_anahtari(x) -> str:
    """'15.150.1005', '15 150 1005', 'Y.15.150/1005' gibi yazımları eşitler."""
    return re.sub(r"[^0-9a-z]", "", katla(x))


def tablo_oku(yol: Path) -> list[list]:
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, data_only=True, read_only=True)
        try:
            satirlar = [list(r) for r in wb.active.iter_rows(values_only=True)]
        finally:
            wb.close()
    else:
        metin = None
        for kod in ("utf-8-sig", "cp1254"):
            try:
                metin = yol.read_text(encoding=kod)
                break
            except UnicodeDecodeError:
                continue
        ilk = "\n".join(metin.splitlines()[:10])
        satirlar = list(csv.reader(metin.splitlines(), delimiter=max(";\t,", key=ilk.count)))
    satirlar = [r for r in satirlar if any(c not in (None, "") for c in r)]
    bi = next((i for i, r in enumerate(satirlar[:10]) if any(katla(c) in ("poz no", "poz", "poz numarasi", "is kalemi no") for c in r)), 0)
    return satirlar[bi:]


def _bul(baslik, *adlar):
    b = [katla(x) for x in baslik]
    return next((b.index(katla(a)) for a in adlar if katla(a) in b), None)


def _al(r, i):
    return r[i] if i is not None and i < len(r) else None


POZ = ("poz no", "poz", "poz numarası", "iş kalemi no")


# ----------------------------------------------------------------------------
# Okuma
# ----------------------------------------------------------------------------

def sozlesme_oku(yol: Path) -> dict[str, dict]:
    s = tablo_oku(yol)
    b = s[0]
    i = {k: _bul(b, *v) for k, v in {"poz": POZ, "tanim": ("tanım", "iş kalemi", "açıklama", "imalat"), "birim": ("birim", "ölçü birimi"),
                                     "fiyat": ("birim fiyat", "birim fiyatı", "sözleşme birim fiyatı"),
                                     "miktar": ("sözleşme miktarı", "miktar", "keşif miktarı")}.items()}
    if i["poz"] is None or i["fiyat"] is None:
        raise SystemExit(f"Sözleşme dosyasında Poz No ve Birim Fiyat gerekli. Başlıklar: {b}")
    return {poz_anahtari(_al(r, i["poz"])): {"poz": str(_al(r, i["poz"])).strip(), "tanim": str(_al(r, i["tanim"]) or "").strip(),
                                             "birim": str(_al(r, i["birim"]) or "").strip(), "fiyat": para(_al(r, i["fiyat"])),
                                             "miktar": para(_al(r, i["miktar"])) if i["miktar"] is not None else None}
            for r in s[1:] if _al(r, i["poz"])}


def hakedis_oku(yol: Path) -> list[dict]:
    s = tablo_oku(yol)
    b = s[0]
    i = {k: _bul(b, *v) for k, v in {
        "poz": POZ, "tanim": ("tanım", "iş kalemi", "açıklama", "imalat"), "birim": ("birim", "ölçü birimi"),
        "fiyat": ("birim fiyat", "birim fiyatı"), "onceki": ("önceki miktar", "önceki", "önceki hakediş miktarı", "önceki toplam"),
        "bu": ("bu dönem miktar", "bu dönem", "bu hakediş miktarı", "dönem miktarı"),
        "toplam": ("toplam miktar", "kümülatif miktar", "toplam", "kümülatif"), "tutar": ("toplam tutar", "tutar", "kümülatif tutar")}.items()}
    if i["poz"] is None or (i["toplam"] is None and i["bu"] is None):
        raise SystemExit(f"Hakediş dosyasında Poz No ve Toplam (veya Bu Dönem) Miktar gerekli. Başlıklar: {b}")
    sonuc = []
    for n, r in enumerate(s[1:], 2):
        if not _al(r, i["poz"]) or katla(_al(r, i["poz"])).startswith(("toplam", "genel")):
            continue
        sonuc.append({"satir": n, "poz": str(_al(r, i["poz"])).strip(), "anahtar": poz_anahtari(_al(r, i["poz"])),
                      "tanim": str(_al(r, i["tanim"]) or "").strip(), "birim": str(_al(r, i["birim"]) or "").strip(),
                      "fiyat": para(_al(r, i["fiyat"])) if i["fiyat"] is not None else None,
                      "onceki": para(_al(r, i["onceki"])) if i["onceki"] is not None else None,
                      "bu": para(_al(r, i["bu"])) if i["bu"] is not None else None,
                      "toplam": para(_al(r, i["toplam"])) if i["toplam"] is not None else None,
                      "tutar": para(_al(r, i["tutar"])) if i["tutar"] is not None else None})
    return sonuc


def miktar_oku(yol: Path | None, *sutunlar) -> dict[str, Decimal]:
    if not yol:
        return {}
    s = tablo_oku(yol)
    i_p, i_m = _bul(s[0], *POZ), _bul(s[0], *sutunlar)
    if i_p is None or i_m is None:
        raise SystemExit(f"{yol.name}: Poz No ve miktar sütunu gerekli. Başlıklar: {s[0]}")
    sonuc: dict[str, Decimal] = {}
    for r in s[1:]:
        if _al(r, i_p):
            k = poz_anahtari(_al(r, i_p))
            sonuc[k] = sonuc.get(k, SIFIR) + para(_al(r, i_m))
    return sonuc


# ----------------------------------------------------------------------------
# Kontrol
# ----------------------------------------------------------------------------

def kontrol(sozlesme, hakedis, saha, onceki_onay, tolerans: Decimal, sozlesme_siniri: bool) -> tuple[list[dict], list[dict]]:
    bulgular, satirlar = [], []

    def ekle(seviye, poz, kontrol_ad, aciklama, fark_tutar=None):
        bulgular.append({"seviye": seviye, "poz": poz, "kontrol": kontrol_ad, "aciklama": aciklama, "fark": fark_tutar})

    gorulen = set()
    for h in hakedis:
        k, poz = h["anahtar"], h["poz"]
        if k in gorulen:
            ekle("Hata", poz, "Mükerrer poz", "Aynı poz hakedişte birden fazla satırda")
        gorulen.add(k)
        sz = sozlesme.get(k)
        toplam = h["toplam"] if h["toplam"] is not None else (h["onceki"] or SIFIR) + (h["bu"] or SIFIR)
        if h["onceki"] is not None and h["bu"] is not None and h["toplam"] is not None and h["onceki"] + h["bu"] != h["toplam"]:
            ekle("Hata", poz, "Aritmetik", f"önceki {h['onceki']:g} + bu dönem {h['bu']:g} ≠ toplam {h['toplam']:g}")
        if h["bu"] is not None and h["bu"] < 0:
            ekle("Dikkat", poz, "Negatif dönem", f"bu dönem miktarı {h['bu']:g} (önceki hakedişte fazla ödeme düzeltmesi mi?)")
        if h["fiyat"] is not None and h["tutar"] is not None and abs(yuvarla(toplam * h["fiyat"]) - h["tutar"]) > Decimal("0.05"):
            ekle("Hata", poz, "Tutar hesabı", f"{toplam:g} × {tl(h['fiyat'])} = {tl(yuvarla(toplam * h['fiyat']))}, taşeron {tl(h['tutar'])} yazmış",
                 h["tutar"] - yuvarla(toplam * h["fiyat"]))
        if sz is None:
            ekle("Yüksek", poz, "Sözleşmede yok", "Poz sözleşme birim fiyat cetvelinde yok: ek iş / yeni birim fiyat onayı gerekir",
                 yuvarla(toplam * (h["fiyat"] or SIFIR)))
            fiyat = SIFIR
        else:
            fiyat = sz["fiyat"]
            if h["fiyat"] is not None and h["fiyat"] != sz["fiyat"]:
                ekle("Yüksek", poz, "Birim fiyat", f"taşeron {tl(h['fiyat'])}, sözleşme {tl(sz['fiyat'])}",
                     yuvarla(toplam * (h["fiyat"] - sz["fiyat"])))
            if h["birim"] and sz["birim"] and katla(h["birim"]) != katla(sz["birim"]):
                ekle("Hata", poz, "Birim", f"taşeron '{h['birim']}', sözleşme '{sz['birim']}'")
            if sz["miktar"] and toplam > sz["miktar"]:
                ekle("Dikkat", poz, "Sözleşme miktarı aşımı",
                     f"toplam {toplam:g} > sözleşme {sz['miktar']:g} (%{(toplam / sz['miktar'] - 1) * 100:.1f} artış)".replace(".", ","))
        if onceki_onay and h["onceki"] is not None:
            oo = onceki_onay.get(k, SIFIR)
            if h["onceki"] != oo:
                ekle("Hata", poz, "Önceki miktar", f"taşeron önceki {h['onceki']:g}, onaylı önceki {oo:g}",
                     yuvarla((h["onceki"] - oo) * fiyat))
        sm = saha.get(k) if saha else None
        if saha and sm is None:
            ekle("Dikkat", poz, "Saha metrajı yok", "Bu poz için saha (yeşil defter) metrajı yok; ölçülmeden onaylanmamalı")
        elif sm is not None and toplam > sm * (1 + tolerans / 100):
            ekle("Yüksek", poz, "Saha metrajı aşımı", f"talep {toplam:g}, saha {sm:g} (tolerans %{float(tolerans):g})",
                 yuvarla((toplam - sm) * fiyat))
        # Onaylanabilir
        onay_kum = toplam
        if sm is not None:
            onay_kum = min(onay_kum, sm)
        elif saha:
            onay_kum = onceki_onay.get(k, SIFIR) if onceki_onay else SIFIR
        if sozlesme_siniri and sz and sz["miktar"]:
            onay_kum = min(onay_kum, sz["miktar"])
        if sz is None:
            onay_kum = onceki_onay.get(k, SIFIR) if onceki_onay else SIFIR
        oo = onceki_onay.get(k, SIFIR) if onceki_onay else (h["onceki"] or SIFIR)
        if h["bu"] is not None and sz is not None:                     # talep edilenden fazlası onaylanmaz
            onay_kum = min(onay_kum, oo + h["bu"])
        talep_bu = toplam - (h["onceki"] if h["onceki"] is not None else oo)
        talep_tutar = yuvarla(talep_bu * (h["fiyat"] if h["fiyat"] is not None else fiyat))
        onay_bu = onay_kum - oo
        onay_tutar = yuvarla(onay_bu * fiyat)
        satirlar.append({"poz": poz, "tanim": h["tanim"] or (sz or {}).get("tanim", ""), "birim": h["birim"] or (sz or {}).get("birim", ""),
                         "sozlesme_fiyat": sz["fiyat"] if sz else None, "taseron_fiyat": h["fiyat"], "sozlesme_miktar": sz["miktar"] if sz else None,
                         "saha": sm, "onceki_onay": oo, "talep_toplam": toplam, "talep_bu": talep_bu, "talep_tutar": talep_tutar,
                         "onay_toplam": onay_kum, "onay_bu": onay_bu, "onay_tutar": onay_tutar, "fark": talep_tutar - onay_tutar})
    # Sahada ölçülmüş ama hakedişe konmamış pozlar (bilgi)
    if saha:
        for k, sm in saha.items():
            if k not in gorulen and sm > (onceki_onay.get(k, SIFIR) if onceki_onay else SIFIR):
                ekle("Bilgi", sozlesme.get(k, {}).get("poz", k), "Hakedişte yok", f"sahada {sm:g} ölçülmüş, taşeron bu pozu talep etmemiş")
    return bulgular, satirlar


SEVIYE_SIRA = {"Hata": 0, "Yüksek": 1, "Dikkat": 2, "Bilgi": 3}


def calistir(sozlesme_yolu: Path, hakedis_yolu: Path, cikti: Path, saha_yolu: Path | None = None, onceki_yolu: Path | None = None,
             tolerans: float = 0.0, sozlesme_siniri: bool = False) -> dict:
    sozlesme = sozlesme_oku(sozlesme_yolu)
    hakedis = hakedis_oku(hakedis_yolu)
    saha = miktar_oku(saha_yolu, "kümülatif miktar", "saha metrajı", "toplam miktar", "miktar", "metraj")
    onceki = miktar_oku(onceki_yolu, "onaylanan kümülatif miktar", "onaylanan kümülatif", "onaylı toplam", "kümülatif miktar",
                        "toplam miktar", "miktar")
    bulgular, satirlar = kontrol(sozlesme, hakedis, saha, onceki, Decimal(str(tolerans)), sozlesme_siniri)
    bulgular.sort(key=lambda b: (SEVIYE_SIRA[b["seviye"]], b["poz"]))
    ozet = {"talep": sum((s["talep_tutar"] for s in satirlar), SIFIR), "onay": sum((s["onay_tutar"] for s in satirlar), SIFIR)}
    ozet["fark"] = ozet["talep"] - ozet["onay"]
    _rapor(bulgular, satirlar, ozet, tolerans, sozlesme_siniri, bool(saha), bool(onceki), cikti)
    return {"bulgular": bulgular, "satirlar": satirlar, "ozet": ozet}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
RENK = {"Hata": "F8C9C6", "Yüksek": "FDE2E1", "Dikkat": "FFF4CE", "Bilgi": "E8F0FE"}
PARA = "#,##0.00"
MIK = "#,##0.###"


def _baslik(ws, satir=1):
    for c in ws[satir]:
        c.fill, c.font = BASLIK_DOLGU, BASLIK_YAZI


def _f(x):
    return None if x is None else float(x)


def _rapor(bulgular, satirlar, ozet, tolerans, sozlesme_siniri, sahali, oncekili, cikti):
    wb = Workbook()
    o = wb.active
    o.title = "Özet"
    o.append(["Taşeron hakediş kontrolü"])
    o["A1"].font = Font(bold=True, size=12)
    o.append(["Taşeronun bu dönem talebi", _f(ozet["talep"])])
    o.append(["Onaylanabilir bu dönem tutarı", _f(ozet["onay"])])
    o.append(["Fark (talep − onay)", _f(ozet["fark"])])
    for r in (2, 3, 4):
        o.cell(r, 2).number_format = PARA
    o.append([])
    for s in SEVIYE_SIRA:
        o.append([s, sum(1 for b in bulgular if b["seviye"] == s)])
        o.cell(o.max_row, 1).fill = PatternFill("solid", fgColor=RENK[s])
    o.append([])
    o.append(["Onay kuralı", "kümülatif miktar en çok saha metrajı kadar" if sahali else "saha metrajı verilmedi — taşeron miktarı esas alındı"])
    if sozlesme_siniri:
        o.append(["", "ve en çok sözleşme miktarı kadar"])
    o.append(["", "birim fiyat sözleşmeden; sözleşmede olmayan pozlar onaylanmaz (ek iş/yeni fiyat onayı gerekir)"])
    o.append(["", "bu dönem = onaylanan kümülatif − " + ("onaylı önceki" if oncekili else "taşeronun önceki miktarı")])
    o.append(["Sonraki adım", "Onaylanan bu dönem tutarıyla KDV, tevkifat, stopaj, teminat ve avans kesintilerini 'Hakediş Hesaplama' "
                              "paketinde hesaplayın; vergi uygulamasını mali müşavirinizle teyit edin."])
    o.column_dimensions["A"].width = 30
    o.column_dimensions["B"].width = 110

    b = wb.create_sheet("Bulgular")
    b.append(["Seviye", "Poz No", "Kontrol", "Açıklama", "Tutar Etkisi", "Taşeron Açıklaması / Karar"])
    _baslik(b)
    for x in bulgular:
        b.append([x["seviye"], x["poz"], x["kontrol"], x["aciklama"], _f(x["fark"]), None])
        b.cell(b.max_row, 1).fill = PatternFill("solid", fgColor=RENK[x["seviye"]])
        b.cell(b.max_row, 5).number_format = PARA
    for j, w in enumerate((9, 16, 22, 70, 14, 30), 1):
        b.column_dimensions[get_column_letter(j)].width = w
    b.freeze_panes = "C2"
    b.auto_filter.ref = b.dimensions

    p = wb.create_sheet("Poz Karşılaştırma")
    p.append(["Poz No", "Tanım", "Birim", "Sözleşme Fiyatı", "Taşeron Fiyatı", "Sözleşme Miktarı", "Saha Metrajı", "Onaylı Önceki",
              "Talep Toplam", "Talep Bu Dönem", "Talep Tutarı", "Onay Toplam", "Onay Bu Dönem", "Onay Tutarı", "Fark"])
    _baslik(p)
    for s in satirlar:
        p.append([s["poz"], s["tanim"], s["birim"], _f(s["sozlesme_fiyat"]), _f(s["taseron_fiyat"]), _f(s["sozlesme_miktar"]), _f(s["saha"]),
                  _f(s["onceki_onay"]), _f(s["talep_toplam"]), _f(s["talep_bu"]), _f(s["talep_tutar"]), _f(s["onay_toplam"]),
                  _f(s["onay_bu"]), _f(s["onay_tutar"]), _f(s["fark"])])
        n = p.max_row
        for c in (4, 5, 11, 14, 15):
            p.cell(n, c).number_format = PARA
        for c in (6, 7, 8, 9, 10, 12, 13):
            p.cell(n, c).number_format = MIK
        if s["fark"]:
            p.cell(n, 15).fill = PatternFill("solid", fgColor=RENK["Yüksek"])
    p.append(["TOPLAM", "", "", None, None, None, None, None, None, None, _f(ozet["talep"]), None, None, _f(ozet["onay"]), _f(ozet["fark"])])
    for c in p[p.max_row]:
        c.font = Font(bold=True)
    for c in (11, 14, 15):
        p.cell(p.max_row, c).number_format = PARA
    for j, w in enumerate((16, 36, 7, 12, 12, 11, 11, 11, 11, 11, 14, 11, 11, 14, 13), 1):
        p.column_dimensions[get_column_letter(j)].width = w
    p.freeze_panes = "C2"
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ornek = BURASI / "ornek_veri"
    ap = argparse.ArgumentParser(description="Taşeron hakedişini sözleşme birim fiyatları, saha metrajı ve önceki onaylı hakedişle karşılaştırır.")
    ap.add_argument("--sozlesme", type=Path, default=ornek / "sozlesme_birim_fiyatlar.csv", help="Poz No, Tanım, Birim, Birim Fiyat, Sözleşme Miktarı")
    ap.add_argument("--hakedis", type=Path, default=ornek / "taseron_hakedis_5.csv",
                    help="Taşeron hakedişi: Poz No, Tanım, Birim, Birim Fiyat, Önceki Miktar, Bu Dönem Miktar, Toplam Miktar, Toplam Tutar")
    ap.add_argument("--saha", type=Path, help="Saha (yeşil defter) metrajı: Poz No, Kümülatif Miktar")
    ap.add_argument("--onceki", type=Path, help="Önceki onaylı hakediş: Poz No, Onaylanan Kümülatif Miktar")
    ap.add_argument("--tolerans", type=float, default=0.0, help="Saha metrajı aşımında kabul edilen pay, %% (varsayılan 0)")
    ap.add_argument("--sozlesme-siniri", action="store_true", help="Onaylanan kümülatifi sözleşme miktarıyla da sınırla")
    ap.add_argument("--cikti", type=Path, default=Path("cikti") / "taseron_hakedis_kontrolu.xlsx")
    a = ap.parse_args(argv)
    if a.hakedis == ornek / "taseron_hakedis_5.csv":
        a.saha = a.saha or ornek / "saha_metraji.csv"
        a.onceki = a.onceki or ornek / "onayli_hakedis_4.csv"
    s = calistir(a.sozlesme, a.hakedis, a.cikti, a.saha, a.onceki, a.tolerans, a.sozlesme_siniri)
    from collections import Counter
    say = Counter(b["seviye"] for b in s["bulgular"])
    print(f"[OK] talep {tl(s['ozet']['talep'])} · onaylanabilir {tl(s['ozet']['onay'])} · fark {tl(s['ozet']['fark'])} TL · "
          + " · ".join(f"{k}: {say.get(k, 0)}" for k in SEVIYE_SIRA))
    print(f"[OK] Rapor: {a.cikti.resolve()}")


if __name__ == "__main__":
    sys.exit(main())
