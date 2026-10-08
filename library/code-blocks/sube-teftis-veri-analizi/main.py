"""
Şube Teftiş Veri Analizi — Workers / Workless kod bloğu
Bankacılık › Teftiş Kurulu › Müfettiş Yardımcısı

Şube işlem verisini personel listesiyle birlikte tarar ve teftişte incelenecek riskli örüntüleri listeler:
  - Personelin kendi hesabında veya yakınlarının hesabında yaptığı işlemler.
  - Mesai dışı, hafta sonu ve resmî tatil günü yapılan işlemler.
  - Personelin işlem limitini aşan ve onaysız işlemler; kendi işlemini onaylama (dört göz ihlali); onaylayanın
    limitinin de yetmediği işlemler.
  - Sık iptal (personelin iptal oranı şube ortalamasının katı) ve iptalden kısa süre sonra aynı müşteriye farklı
    tutarla yeniden yapılan işlemler.
  - Aynı müşterinin aynı gün eşiğin altında kalan ama toplamı eşiği aşan nakit işlemleri (bölünmüş işlem).
  - Bir personelin işlemlerinin tek müşteride yoğunlaşması.
Bulgular örnek seçimi ve inceleme içindir; tek başına usulsüzlük göstermez.
İnternete bağlanmaz.

Kullanım:
    python main.py                                                     # örnek şube verisi (Eylül 2026)
    python main.py --islemler islemler.xlsx --personel personel.xlsx --mesai 08:30-17:30 --bolunmus-esik 100000
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")

# Genel tatiller (2912 sayılı Kanun) ve dini bayramlar (Diyanet takvimi; her yıl --tatiller ile kontrol edin)
SABIT_TATILLER = {(1, 1), (4, 23), (5, 1), (5, 19), (7, 15), (8, 30), (10, 29)}
DINI_BAYRAMLAR = {2025: [(date(2025, 3, 30), 3), (date(2025, 6, 6), 4)], 2026: [(date(2026, 3, 20), 3), (date(2026, 5, 27), 4)],
                  2027: [(date(2027, 3, 9), 3), (date(2027, 5, 16), 4)]}

SUTUNLAR = {
    "no": ("islem no", "islem numarasi", "fis no", "referans no", "no"),
    "tarih": ("tarih", "islem tarihi"),
    "saat": ("saat", "islem saati", "zaman"),
    "sube": ("sube", "sube kodu", "sube adi"),
    "personel": ("personel", "personel sicil", "sicil", "kullanici", "islemi yapan"),
    "musteri": ("musteri no", "musteri", "hesap sahibi no"),
    "tur": ("islem turu", "tur", "islem tipi", "islem kodu aciklamasi"),
    "tutar": ("tutar", "islem tutari", "tl karsiligi"),
    "durum": ("durum", "islem durumu"),
    "onaylayan": ("onaylayan", "onaylayan sicil", "onay veren", "yetkili"),
    "aciklama": ("aciklama", "not"),
}
PERSONEL_SUTUNLARI = {
    "sicil": ("sicil", "sicil no", "personel", "kullanici"),
    "ad": ("ad soyad", "adi soyadi", "ad"),
    "gorev": ("gorev", "unvan"),
    "sube": ("sube",),
    "limit": ("islem limiti", "limit", "yetki limiti"),
    "musteri": ("musteri no", "kendi musteri no"),
    "yakin": ("yakin musteri no", "yakinlari", "yakin musteri"),
}
ONEM_PUAN = {"Yüksek": 3, "Orta": 2, "Bilgi": 1}


def katla(s) -> str:
    return " ".join(str(s or "").translate(_TR).lower().split())


def tarih(x) -> date | None:
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    for f in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S"):
        try:
            return datetime.strptime(str(x or "").strip(), f).date()
        except ValueError:
            pass
    return None


def saat(x, tarih_hucresi=None) -> time | None:
    if isinstance(x, time):
        return x
    if isinstance(x, datetime):
        return x.time()
    m = re.search(r"(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?", str(x or "")) or re.search(r"\s(\d{1,2}):(\d{2})(?::(\d{2}))?", str(tarih_hucresi or ""))
    if not m:
        return tarih_hucresi.time() if isinstance(tarih_hucresi, datetime) else time(0, 0)
    return time(int(m.group(1)), int(m.group(2)), int(m.group(3) or 0))


def para(x) -> Decimal | None:
    if x in (None, ""):
        return None
    if isinstance(x, (int, float)):
        return Decimal(str(x))
    s = re.sub(r"[^\d,.\-]", "", str(x))
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[1]) == 3):
        s = s.replace(".", "")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def tl(x) -> str:
    return "—" if x is None else f"{x:,.0f}".replace(",", ".")


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
    return [r for r in satirlar if any(c not in (None, "") for c in r)]


def eslestir(satirlar: list[list], sozluk: dict, zorunlu: tuple[str, ...]) -> tuple[int, dict[str, int]]:
    """Başlık satırını (üstteki başlık satırlarını atlayarak) ve alan → sütun eşleşmesini bulur."""
    for i, r in enumerate(satirlar[:15]):
        b = [katla(c) for c in r]
        es = {}
        for alan, adlar in sozluk.items():
            j = next((b.index(a) for a in adlar if a in b), None)
            if j is not None and j not in es.values():
                es[alan] = j
        if all(z in es for z in zorunlu):
            return i, es
    raise ValueError(f"Başlık satırı bulunamadı; gerekli sütunlar: {', '.join(zorunlu)}")


def tatil_kumesi(yillar: set[int], ek: set[date]) -> set[date]:
    t = set(ek)
    for y in yillar:
        t |= {date(y, a, g) for a, g in SABIT_TATILLER}
        for ilk, n in DINI_BAYRAMLAR.get(y, []):
            t |= {ilk + timedelta(days=i) for i in range(n)}
    return t


# ----------------------------------------------------------------------------
# Veri
# ----------------------------------------------------------------------------

@dataclass
class Islem:
    satir: int
    no: str
    tarih: date
    saat: time
    sube: str
    personel: str
    musteri: str
    tur: str
    tutar: Decimal
    iptal: bool
    onaylayan: str
    aciklama: str
    isaretler: list = field(default_factory=list)

    @property
    def an(self) -> datetime:
        return datetime.combine(self.tarih, self.saat)

    @property
    def nakit(self) -> bool:
        return "nakit" in katla(self.tur) or "vezne" in katla(self.tur)


def islemleri_oku(yol: Path) -> tuple[list[Islem], list[str]]:
    s = tablo_oku(yol)
    bi, es = eslestir(s, SUTUNLAR, ("tarih", "personel", "tutar"))
    al = lambda r, a: r[es[a]] if a in es and es[a] < len(r) else None  # noqa: E731
    islemler, hatalar = [], []
    for n, r in enumerate(s[bi + 1:], bi + 2):
        t, tu = tarih(al(r, "tarih")), para(al(r, "tutar"))
        if t is None or tu is None:
            hatalar.append(f"Satır {n}: tarih veya tutar okunamadı")
            continue
        durum = katla(al(r, "durum"))
        islemler.append(Islem(n, str(al(r, "no") or f"S{n}").strip(), t, saat(al(r, "saat"), al(r, "tarih")), str(al(r, "sube") or "").strip(),
                              str(al(r, "personel") or "").strip(), str(al(r, "musteri") or "").strip(), str(al(r, "tur") or "").strip(), abs(tu),
                              "iptal" in durum or "ters" in durum or "red" in durum, str(al(r, "onaylayan") or "").strip(),
                              str(al(r, "aciklama") or "").strip()))
    return islemler, hatalar


def coklu(x) -> set[str]:
    return {p.strip() for p in re.split(r"[|,/;]", str(x or "")) if p.strip()}


def personel_oku(yol: Path | None) -> dict[str, dict]:
    if yol is None or not yol.exists():
        return {}
    s = tablo_oku(yol)
    bi, es = eslestir(s, PERSONEL_SUTUNLARI, ("sicil",))
    al = lambda r, a: r[es[a]] if a in es and es[a] < len(r) else None  # noqa: E731
    return {str(al(r, "sicil")).strip(): {"ad": str(al(r, "ad") or "").strip(), "gorev": str(al(r, "gorev") or "").strip(),
                                          "sube": str(al(r, "sube") or "").strip(), "limit": para(al(r, "limit")),
                                          "musteri": coklu(al(r, "musteri")), "yakin": coklu(al(r, "yakin"))}
            for r in s[bi + 1:] if al(r, "sicil")}


# ----------------------------------------------------------------------------
# Analiz
# ----------------------------------------------------------------------------

@dataclass
class Ayarlar:
    mesai_bas: time = time(8, 30)
    mesai_bit: time = time(17, 30)
    cumartesi_acik: bool = False
    ek_tatiller: set = field(default_factory=set)
    iptal_kat: Decimal = Decimal(2)
    iptal_min_adet: int = 3
    iptal_min_oran: Decimal = Decimal("0.05")
    yeniden_dk: int = 60
    bolunmus_esik: Decimal = Decimal(100000)
    yogunlasma_adet: int = 5
    yogunlasma_oran: Decimal = Decimal("0.10")


def analiz_et(islemler: list[Islem], personel: dict[str, dict], ay: Ayarlar) -> dict:
    bulgular = []

    def b(onem, tur, aciklama, personel_="", musteri="", ilgili=()):
        bulgular.append({"onem": onem, "tur": tur, "personel": personel_, "musteri": musteri, "islemler": [i.no for i in ilgili], "aciklama": aciklama})
        for i in ilgili:
            i.isaretler.append((onem, tur))

    tatiller = tatil_kumesi({i.tarih.year for i in islemler}, ay.ek_tatiller)
    musteri_personel = {m: s for s, p in personel.items() for m in p["musteri"]}
    yakin_personel = {m: s for s, p in personel.items() for m in p["yakin"]}
    tamam = [i for i in islemler if not i.iptal]

    for s in sorted({i.personel for i in islemler} - set(personel)):
        if personel:
            b("Bilgi", "Tanımsız personel", f"{s} personel listesinde yok; limit ve ilişki kontrolleri yapılamadı", s,
              ilgili=[i for i in islemler if i.personel == s][:1])

    for i in tamam:
        p = personel.get(i.personel, {})
        if i.musteri and i.musteri in p.get("musteri", ()):
            b("Yüksek", "Personelin kendi hesabı", f"{i.personel} kendi müşteri numarasında ({i.musteri}) {i.tur} {tl(i.tutar)} TL işlem yapmış",
              i.personel, i.musteri, [i])
        elif i.musteri and i.musteri in p.get("yakin", ()):
            b("Orta", "Personel yakınının hesabı", f"{i.personel} yakınının hesabında ({i.musteri}) {i.tur} {tl(i.tutar)} TL işlem yapmış",
              i.personel, i.musteri, [i])
        elif i.musteri in musteri_personel and musteri_personel[i.musteri] != i.personel:
            b("Bilgi", "Başka personelin hesabı", f"{i.personel}, {musteri_personel[i.musteri]} sicilli personelin hesabında işlem yapmış",
              i.personel, i.musteri, [i])
        elif i.musteri in yakin_personel and yakin_personel[i.musteri] != i.personel:
            b("Bilgi", "Personel yakınının hesabı (başka personel)", f"{i.musteri}, {yakin_personel[i.musteri]} sicilli personelin yakını; işlemi {i.personel} yapmış",
              i.personel, i.musteri, [i])

    for i in tamam:
        tatil = i.tarih in tatiller
        hafta_sonu = i.tarih.weekday() == 6 or (i.tarih.weekday() == 5 and not ay.cumartesi_acik)
        if tatil or hafta_sonu:
            b("Orta", "Tatil / hafta sonu işlemi", f"{i.tarih:%d.%m.%Y} ({'resmî tatil' if tatil else 'hafta sonu'}) {i.saat:%H:%M} {i.tur} {tl(i.tutar)} TL",
              i.personel, i.musteri, [i])
        elif not (ay.mesai_bas <= i.saat <= ay.mesai_bit):
            b("Orta", "Mesai dışı işlem", f"{i.tarih:%d.%m.%Y} {i.saat:%H:%M} {i.tur} {tl(i.tutar)} TL (mesai {ay.mesai_bas:%H:%M}–{ay.mesai_bit:%H:%M})",
              i.personel, i.musteri, [i])

    for i in tamam:
        p, o = personel.get(i.personel), personel.get(i.onaylayan)
        if i.onaylayan and i.onaylayan == i.personel:
            b("Yüksek", "Kendi işlemini onaylama", f"{i.personel} {tl(i.tutar)} TL tutarlı {i.tur} işlemini kendisi onaylamış (dört göz ilkesi)",
              i.personel, i.musteri, [i])
        elif p and p["limit"] is not None and i.tutar > p["limit"]:
            if not i.onaylayan:
                b("Yüksek", "Limit aşımı — onay yok", f"{tl(i.tutar)} TL {i.tur}, personel limiti {tl(p['limit'])} TL; onaylayan yok", i.personel, i.musteri, [i])
            elif o and o["limit"] is not None and i.tutar > o["limit"]:
                b("Orta", "Onaylayanın limiti yetersiz", f"{tl(i.tutar)} TL {i.tur}, onaylayan {i.onaylayan} limiti {tl(o['limit'])} TL", i.personel, i.musteri, [i])

    toplam_oran = Decimal(sum(i.iptal for i in islemler)) / len(islemler) if islemler else Decimal(0)
    per_islem = defaultdict(list)
    for i in islemler:
        per_islem[i.personel].append(i)
    for s, lst in per_islem.items():
        ipt = [i for i in lst if i.iptal]
        oran = Decimal(len(ipt)) / len(lst)
        if len(ipt) >= ay.iptal_min_adet and oran >= ay.iptal_min_oran and oran >= ay.iptal_kat * toplam_oran:
            b("Orta", "Sık iptal", f"{s}: {len(lst)} işlemden {len(ipt)} iptal (%{oran * 100:.1f}); şube geneli %{toplam_oran * 100:.1f}".replace(".", ","),
              s, ilgili=ipt)
    for i in islemler:
        if not i.iptal:
            continue
        sonra = [j for j in tamam if j.personel == i.personel and j.musteri == i.musteri and katla(j.tur) == katla(i.tur)
                 and timedelta(0) <= j.an - i.an <= timedelta(minutes=ay.yeniden_dk) and j.tutar != i.tutar]
        for j in sonra[:1]:
            dk = int((j.an - i.an).total_seconds() // 60)
            b("Orta", "İptal sonrası farklı tutarla yeniden işlem", f"{i.no} ({tl(i.tutar)} TL) iptal edildikten {dk} dk sonra {j.no} ({tl(j.tutar)} TL) yapılmış",
              i.personel, i.musteri, [i, j])

    gun = defaultdict(list)
    for i in tamam:
        if i.nakit and i.musteri:
            gun[(i.musteri, i.tarih)].append(i)
    for (m, t), lst in sorted(gun.items(), key=lambda x: (x[0][1], x[0][0])):
        top = sum((i.tutar for i in lst), Decimal(0))
        if len(lst) >= 2 and all(i.tutar < ay.bolunmus_esik for i in lst) and top >= ay.bolunmus_esik:
            b("Orta", "Bölünmüş nakit işlem", f"{m}: {t:%d.%m.%Y} günü {len(lst)} nakit işlem, her biri {tl(ay.bolunmus_esik)} TL'nin altında, toplam {tl(top)} TL "
              f"(personel: {', '.join(sorted({i.personel for i in lst}))})", ", ".join(sorted({i.personel for i in lst})), m, lst)

    for s, lst in per_islem.items():
        say = Counter(i.musteri for i in lst if i.musteri and not i.iptal)
        for m, n in say.most_common():
            if n >= ay.yogunlasma_adet and Decimal(n) / len(lst) >= ay.yogunlasma_oran:
                b("Bilgi", "Personel-müşteri yoğunlaşması", f"{s}: {len(lst)} işlemden {n} tanesi (%{Decimal(n) / len(lst) * 100:.1f}) {m} numaralı müşteride".replace(".", ","),
                  s, m, [i for i in lst if i.musteri == m and not i.iptal])

    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    bulgular.sort(key=lambda x: (sira[x["onem"]], x["tur"], x["islemler"][:1]))
    ozet = []
    for s in sorted(per_islem):
        lst = per_islem[s]
        ilgili = [x for x in bulgular if s in x["personel"].split(", ")]
        ozet.append({"sicil": s, "ad": personel.get(s, {}).get("ad", ""), "gorev": personel.get(s, {}).get("gorev", ""), "islem": len(lst),
                     "iptal": sum(i.iptal for i in lst), "tutar": sum((i.tutar for i in lst if not i.iptal), Decimal(0)),
                     "turler": Counter(x["tur"] for x in ilgili), "puan": sum(ONEM_PUAN[x["onem"]] for x in ilgili)})
    ozet.sort(key=lambda x: (-x["puan"], x["sicil"]))
    return {"bulgular": bulgular, "ozet": ozet, "iptal_orani": toplam_oran, "tatiller": tatiller}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
RENK = {"Yüksek": "FDE2E1", "Orta": "FFF4CE", "Bilgi": "E8F0FE"}
KONTROL = PatternFill("solid", fgColor="FFF4CE")
UST = Alignment(vertical="top", wrap_text=True)


def _baslik(ws, basliklar, genislik):
    ws.append(basliklar)
    for h in ws[ws.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def rapor_yaz(cikti: Path, islemler: list[Islem], personel: dict, s: dict, ay: Ayarlar, hatalar: list[str]) -> None:
    wb = Workbook()
    o = wb.active
    o.title = "Özet"
    o.column_dimensions["A"].width, o.column_dimensions["B"].width = 34, 70
    tarihler = [i.tarih for i in islemler]
    say = Counter(x["onem"] for x in s["bulgular"])
    for a, d in [("Dönem", f"{min(tarihler):%d.%m.%Y} – {max(tarihler):%d.%m.%Y}" if tarihler else "—"),
                 ("İşlem sayısı", len(islemler)), ("Personel sayısı", len({i.personel for i in islemler})),
                 ("Şube iptal oranı", f"%{s['iptal_orani'] * 100:.1f}".replace(".", ",")),
                 ("Bulgu: Yüksek / Orta / Bilgi", f"{say['Yüksek']} / {say['Orta']} / {say['Bilgi']}"),
                 ("İşaretli işlem", sum(1 for i in islemler if i.isaretler)),
                 ("Okunamayan satır", "; ".join(hatalar) or "—")]:
        o.append([a, d])
        o.cell(o.max_row, 1).font = Font(bold=True)
    o.append([])
    o.append(["Bulgu türü", "Adet"])
    for h in o[o.max_row]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
    for t, n in Counter(x["tur"] for x in s["bulgular"]).most_common():
        o.append([t, n])
    o.append([])
    o.append(["Not", "Bulgular inceleme ve örnek seçimi içindir; tek başına usulsüzlük göstermez. Açıklamaları işlem belgeleri ve "
                     "ilgili personelle görüşerek değerlendirin."])
    o.cell(o.max_row, 2).alignment = UST

    bl = wb.create_sheet("Bulgular")
    _baslik(bl, ["Önem", "Tür", "Personel", "Müşteri No", "İşlemler", "Açıklama", "Müfettiş Değerlendirmesi"], (9, 30, 12, 12, 24, 80, 30))
    for x in s["bulgular"]:
        bl.append([x["onem"], x["tur"], x["personel"], x["musteri"], ", ".join(x["islemler"][:12]) + (" …" if len(x["islemler"]) > 12 else ""), x["aciklama"], ""])
        bl.cell(bl.max_row, 1).fill = PatternFill("solid", fgColor=RENK[x["onem"]])
        bl.cell(bl.max_row, 7).fill = KONTROL
        for h in bl[bl.max_row]:
            h.alignment = UST
    bl.auto_filter.ref = f"A1:G{bl.max_row}"

    turler = sorted({x["tur"] for x in s["bulgular"]})
    po = wb.create_sheet("Personel Özeti")
    _baslik(po, ["Sicil", "Ad Soyad", "Görev", "İşlem", "İptal", "İptal %", "Tutar (iptal hariç)", "Risk Puanı", *turler], (9, 18, 20, 8, 7, 8, 16, 9, *[14] * len(turler)))
    for x in s["ozet"]:
        po.append([x["sicil"], x["ad"], x["gorev"], x["islem"], x["iptal"], x["iptal"] / x["islem"] if x["islem"] else 0, float(x["tutar"]), x["puan"],
                   *[x["turler"].get(t) or None for t in turler]])
        po.cell(po.max_row, 6).number_format = "0.0%"
        po.cell(po.max_row, 7).number_format = "#,##0"
        if x["puan"] >= 5:
            po.cell(po.max_row, 8).fill = PatternFill("solid", fgColor=RENK["Yüksek"])

    ws = wb.create_sheet("İşlemler")
    _baslik(ws, ["İşlem No", "Tarih", "Saat", "Şube", "Personel", "Müşteri No", "İşlem Türü", "Tutar", "Durum", "Onaylayan", "Açıklama", "İşaretler"],
            (10, 11, 7, 14, 9, 11, 18, 13, 11, 10, 24, 50))
    for i in islemler:
        ws.append([i.no, i.tarih, i.saat.strftime("%H:%M"), i.sube, i.personel, i.musteri, i.tur, float(i.tutar), "İptal" if i.iptal else "Tamamlandı",
                   i.onaylayan, i.aciklama, "; ".join(dict.fromkeys(t for _, t in i.isaretler))])
        ws.cell(ws.max_row, 2).number_format = "DD.MM.YYYY"
        ws.cell(ws.max_row, 8).number_format = "#,##0.00"
        if i.isaretler:
            en = max((o_ for o_, _ in i.isaretler), key=ONEM_PUAN.get)
            ws.cell(ws.max_row, 12).fill = PatternFill("solid", fgColor=RENK[en])
    ws.auto_filter.ref = f"A1:L{ws.max_row}"

    pa = wb.create_sheet("Parametreler")
    _baslik(pa, ["Parametre", "Değer"], (40, 50))
    for a, d in [("Mesai", f"{ay.mesai_bas:%H:%M} – {ay.mesai_bit:%H:%M}"), ("Cumartesi açık", "Evet" if ay.cumartesi_acik else "Hayır"),
                 ("Sık iptal", f"en az {ay.iptal_min_adet} iptal, oran ≥ %{ay.iptal_min_oran * 100:g} ve şube oranının ≥ {ay.iptal_kat:g} katı"),
                 ("İptal sonrası yeniden işlem penceresi", f"{ay.yeniden_dk} dk"), ("Bölünmüş nakit eşiği", f"{tl(ay.bolunmus_esik)} TL (müşteri × gün)"),
                 ("Yoğunlaşma", f"en az {ay.yogunlasma_adet} işlem ve personel işlemlerinin ≥ %{ay.yogunlasma_oran * 100:g}'i"),
                 ("Personel listesi", f"{len(personel)} kişi" if personel else "verilmedi (ilişki ve limit kontrolleri yapılmadı)")]:
        pa.append([a, d])
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(islem_yolu: Path, personel_yolu: Path | None, cikti: Path, ay: Ayarlar | None = None) -> dict:
    ay = ay or Ayarlar()
    islemler, hatalar = islemleri_oku(islem_yolu)
    personel = personel_oku(personel_yolu)
    s = analiz_et(islemler, personel, ay)
    rapor_yaz(cikti, islemler, personel, s, ay, hatalar)
    return {**s, "islemler": islemler, "personel": personel, "hatalar": hatalar}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="Şube işlem verisinde teftiş için riskli örüntüleri tarar.")
    p.add_argument("--islemler", type=Path, default=ORNEK / "islemler.csv", help="İşlem listesi (.xlsx/.csv)")
    p.add_argument("--personel", type=Path, default=ORNEK / "personel.csv", help="Personel listesi: sicil, limit, müşteri no, yakın müşteri no")
    p.add_argument("--mesai", default="08:30-17:30", help="Mesai saatleri (varsayılan 08:30-17:30)")
    p.add_argument("--cumartesi", action="store_true", help="Şube cumartesi açık (cumartesi işlemleri hafta sonu sayılmaz)")
    p.add_argument("--tatiller", type=Path, help="Ek tatil günleri (.csv/.txt, satır başına GG.AA.YYYY)")
    p.add_argument("--bolunmus-esik", default="100000", help="Bölünmüş nakit işlem eşiği, TL (varsayılan 100000; kurum eşiğinizi girin)")
    p.add_argument("--iptal-kat", default="2", help="Personel iptal oranı şube oranının kaç katıysa (varsayılan 2)")
    p.add_argument("--yeniden-dk", type=int, default=60, help="İptalden sonra yeniden işlem penceresi, dakika (varsayılan 60)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "sube_teftis_analizi.xlsx")
    a = p.parse_args(argv)
    m = re.fullmatch(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})", a.mesai.strip())
    if not m:
        print(f"[X] Mesai SS:DD-SS:DD olmalı: {a.mesai}")
        return 2
    ek = set()
    if a.tatiller:
        ek = {d for d in (tarih(x) for x in re.findall(r"\d{1,2}\.\d{1,2}\.\d{4}|\d{4}-\d\d-\d\d", a.tatiller.read_text(encoding="utf-8-sig"))) if d}
    ay = Ayarlar(time(int(m.group(1)), int(m.group(2))), time(int(m.group(3)), int(m.group(4))), a.cumartesi, ek, Decimal(a.iptal_kat),
                 yeniden_dk=a.yeniden_dk, bolunmus_esik=para(a.bolunmus_esik))
    if not a.islemler.exists():
        print(f"[X] Dosya bulunamadı: {a.islemler}")
        return 1
    try:
        s = calistir(a.islemler, a.personel, a.cikti, ay)
    except ValueError as h:
        print(f"[X] {h}")
        return 1
    say = Counter(x["onem"] for x in s["bulgular"])
    print(f"[OK] {len(s['islemler'])} işlem · {len(s['bulgular'])} bulgu (Yüksek {say['Yüksek']}, Orta {say['Orta']}, Bilgi {say['Bilgi']})")
    for x in s["bulgular"]:
        if x["onem"] == "Yüksek":
            print(f"[X] {x['tur']}: {x['aciklama']}")
    for h in s["hatalar"]:
        print(f"[!] {h}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
