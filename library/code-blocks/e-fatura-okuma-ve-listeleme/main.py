"""
e-Fatura Okuma ve Listeleme — Workers / Workless kod bloğu
Muhasebe › Ön Muhasebe Elemanı (ayrıca: Mali Müşavirlik Bürosu › Muhasebe Elemanı)

GİB UBL-TR biçimindeki e-Fatura ve e-Arşiv XML dosyalarını (tek tek veya ZIP içinde) okur;
fatura ve satır bazında Excel listesi, KDV oranı kırılımı, tevkifat ve diğer vergiler ile
tutarlılık kontrollerini üretir. İnternete bağlanmaz.

Kullanım:
    python main.py                                         # örnek faturalarla dener
    python main.py --girdi "C:/Faturalar/Eylul" --cikti cikti/faturalar.xlsx
"""
from __future__ import annotations

import argparse
import io
import sys
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import kimlik_cekirdek as kimlik

BURASI = Path(__file__).resolve().parent

NS = {
    "inv": "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
}

# Vergi türü kodları (GİB UBL-TR kod listesi; adlar GİB vergi kodu tablosundan)
VERGI_ADLARI = {
    "0003": "Gelir Vergisi Stopajı", "0011": "Kurumlar Vergisi Stopajı",
    "0015": "KDV", "0021": "Banka Muameleleri Vergisi", "0022": "Sigorta Muameleleri Vergisi",
    "0061": "KKDF Kesintisi", "0071": "ÖTV (Petrol ve Doğalgaz Ürünleri)",
    "0073": "ÖTV (Kolalı Gazoz, Alkollü İçecek ve Tütün)", "0074": "ÖTV (Dayanıklı Tüketim ve Diğer Mallar)",
    "0075": "ÖTV (Alkollü İçecekler)", "0076": "ÖTV (Tütün Mamulleri)", "0077": "ÖTV (Kolalı Gazozlar)",
    "1047": "Damga Vergisi", "1048": "5035 SK Damga Vergisi", "4071": "Elektrik ve Havagazı Tüketim Vergisi",
    "4080": "Özel İletişim Vergisi", "4081": "5035 SK Özel İletişim Vergisi", "4171": "ÖTV Tevkifatı (Petrol ve Doğalgaz)",
    "9015": "KDV Tevkifatı", "9021": "4961 Banka Sigorta Muameleleri Vergisi", "9077": "ÖTV (Tescile Tabi Motorlu Taşıtlar)",
}

# UN/ECE Tavsiye No. 20 birim kodları (yaygın olanlar)
BIRIMLER = {
    "C62": "Adet", "NIU": "Adet", "KGM": "kg", "GRM": "g", "TNE": "ton", "LTR": "lt", "MLT": "ml",
    "MTR": "m", "CMT": "cm", "KMT": "km", "MTK": "m²", "MTQ": "m³", "KWH": "kWh", "MWH": "MWh",
    "SET": "Set", "PA": "Paket", "BX": "Kutu", "PR": "Çift", "DZN": "Düzine",
    "HUR": "Saat", "DAY": "Gün", "MON": "Ay", "ANN": "Yıl", "MIN": "Dakika",
}


def D(x) -> Decimal:
    try:
        return Decimal(str(x).strip()) if x not in (None, "") else Decimal(0)
    except InvalidOperation:
        return Decimal(0)


def metin(el, yol: str, varsayilan: str = "") -> str:
    if el is None:
        return varsayilan
    bulunan = el.find(yol, NS)
    return (bulunan.text or "").strip() if bulunan is not None and bulunan.text else varsayilan


# ----------------------------------------------------------------------------
# Veri yapıları
# ----------------------------------------------------------------------------

@dataclass
class Satir:
    sira: str
    urun: str
    miktar: Decimal
    birim: str
    birim_fiyat: Decimal
    iskonto: Decimal
    tutar: Decimal
    kdv_orani: Decimal | None
    kdv_tutari: Decimal


@dataclass
class Fatura:
    kaynak: str
    no: str = ""
    ettn: str = ""
    tarih: str = ""
    senaryo: str = ""
    tip: str = ""
    para_birimi: str = "TRY"
    kur: Decimal = Decimal(1)
    satici_vkn: str = ""
    satici_unvan: str = ""
    alici_vkn: str = ""
    alici_unvan: str = ""
    mal_hizmet: Decimal = Decimal(0)
    iskonto: Decimal = Decimal(0)
    vergi_haric: Decimal = Decimal(0)
    vergi_dahil: Decimal = Decimal(0)
    odenecek: Decimal = Decimal(0)
    kdv: dict = field(default_factory=dict)            # oran -> [matrah, kdv]
    diger_vergiler: dict = field(default_factory=dict)  # kod -> tutar
    tevkifat: Decimal = Decimal(0)
    tevkifat_kodlari: list = field(default_factory=list)
    satirlar: list = field(default_factory=list)
    uyarilar: list = field(default_factory=list)

    @property
    def toplam_kdv(self) -> Decimal:
        return sum((v[1] for v in self.kdv.values()), Decimal(0))


# ----------------------------------------------------------------------------
# Okuma
# ----------------------------------------------------------------------------

def taraf(kok, yol: str) -> tuple[str, str]:
    party = kok.find(f"{yol}/cac:Party", NS)
    if party is None:
        return "", ""
    vkn = ""
    for pid in party.findall("cac:PartyIdentification/cbc:ID", NS):
        if pid.get("schemeID") in {"VKN", "TCKN"}:
            vkn = (pid.text or "").strip()
            break
    ad = metin(party, "cac:PartyName/cbc:Name")
    if not ad:
        kisi = party.find("cac:Person", NS)
        ad = " ".join(filter(None, [metin(kisi, "cbc:FirstName"), metin(kisi, "cbc:FamilyName")]))
    return vkn, ad


def fatura_oku(veri: bytes, kaynak: str) -> Fatura:
    f = Fatura(kaynak=kaynak)
    try:
        kok = ET.fromstring(veri)
    except ET.ParseError as h:
        f.uyarilar.append(f"XML okunamadı: {h}")
        return f
    if not kok.tag.endswith("}Invoice"):
        f.uyarilar.append("UBL Invoice belgesi değil (irsaliye veya başka bir belge olabilir)")
        return f

    f.no = metin(kok, "cbc:ID")
    f.ettn = metin(kok, "cbc:UUID")
    f.tarih = metin(kok, "cbc:IssueDate")
    f.senaryo = metin(kok, "cbc:ProfileID")
    f.tip = metin(kok, "cbc:InvoiceTypeCode")
    f.para_birimi = metin(kok, "cbc:DocumentCurrencyCode", "TRY")
    if f.para_birimi != "TRY":
        f.kur = D(metin(kok, "cac:PricingExchangeRate/cbc:CalculationRate")) or Decimal(0)
        if not f.kur:
            f.uyarilar.append("Döviz faturasında kur (PricingExchangeRate) yok")
    f.satici_vkn, f.satici_unvan = taraf(kok, "cac:AccountingSupplierParty")
    f.alici_vkn, f.alici_unvan = taraf(kok, "cac:AccountingCustomerParty")

    lmt = kok.find("cac:LegalMonetaryTotal", NS)
    f.mal_hizmet = D(metin(lmt, "cbc:LineExtensionAmount"))
    f.iskonto = D(metin(lmt, "cbc:AllowanceTotalAmount"))
    f.vergi_haric = D(metin(lmt, "cbc:TaxExclusiveAmount"))
    f.vergi_dahil = D(metin(lmt, "cbc:TaxInclusiveAmount"))
    f.odenecek = D(metin(lmt, "cbc:PayableAmount"))

    for tt in kok.findall("cac:TaxTotal", NS):
        for st in tt.findall("cac:TaxSubtotal", NS):
            kod = metin(st, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode")
            tutar = D(metin(st, "cbc:TaxAmount"))
            if kod == "0015":
                oran = D(metin(st, "cbc:Percent"))
                matrah = D(metin(st, "cbc:TaxableAmount"))
                k = f.kdv.setdefault(oran, [Decimal(0), Decimal(0)])
                k[0] += matrah
                k[1] += tutar
            elif kod:
                f.diger_vergiler[kod] = f.diger_vergiler.get(kod, Decimal(0)) + tutar
    for wt in kok.findall("cac:WithholdingTaxTotal", NS):
        f.tevkifat += D(metin(wt, "cbc:TaxAmount"))
        for st in wt.findall("cac:TaxSubtotal", NS):
            kod = metin(st, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode")
            oran = metin(st, "cbc:Percent")
            if kod:
                f.tevkifat_kodlari.append(f"{kod} (%{oran})" if oran else kod)

    for sat in kok.findall("cac:InvoiceLine", NS):
        miktar_el = sat.find("cbc:InvoicedQuantity", NS)
        kdv_orani, kdv_tutari = None, Decimal(0)
        for st in sat.findall("cac:TaxTotal/cac:TaxSubtotal", NS):
            if metin(st, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode") == "0015":
                kdv_orani = D(metin(st, "cbc:Percent"))
                kdv_tutari += D(metin(st, "cbc:TaxAmount"))
        iskonto = sum((D(metin(ac, "cbc:Amount")) for ac in sat.findall("cac:AllowanceCharge", NS)
                       if metin(ac, "cbc:ChargeIndicator") == "false"), Decimal(0))
        birim = miktar_el.get("unitCode", "") if miktar_el is not None else ""
        f.satirlar.append(Satir(
            sira=metin(sat, "cbc:ID"),
            urun=metin(sat, "cac:Item/cbc:Name"),
            miktar=D(miktar_el.text if miktar_el is not None else 0),
            birim=BIRIMLER.get(birim, birim),
            birim_fiyat=D(metin(sat, "cac:Price/cbc:PriceAmount")),
            iskonto=iskonto,
            tutar=D(metin(sat, "cbc:LineExtensionAmount")),
            kdv_orani=kdv_orani,
            kdv_tutari=kdv_tutari,
        ))
    kontrol_et(f)
    return f


def kontrol_et(f: Fatura) -> None:
    tol = Decimal("0.05")
    satir_toplami = sum((s.tutar for s in f.satirlar), Decimal(0))
    if f.satirlar and abs(satir_toplami - f.mal_hizmet) > tol:
        f.uyarilar.append(f"Satır toplamı ({satir_toplami}) ile mal/hizmet toplamı ({f.mal_hizmet}) farklı")
    for oran, (matrah, kdv) in f.kdv.items():
        beklenen = (matrah * oran / 100).quantize(Decimal("0.01"))
        if abs(beklenen - kdv) > Decimal("0.02") + Decimal("0.01") * len(f.satirlar):
            f.uyarilar.append(f"%{oran:g} KDV: matrah × oran = {beklenen}, faturada {kdv}")
    diger = sum(f.diger_vergiler.values(), Decimal(0))
    if f.vergi_dahil and abs(f.vergi_haric + f.toplam_kdv + diger - f.vergi_dahil) > tol:
        f.uyarilar.append("Vergiler hariç + vergiler ≠ vergiler dahil tutar")
    if f.vergi_dahil and f.odenecek and abs(f.vergi_dahil - f.tevkifat - f.odenecek) > tol:
        f.uyarilar.append("Vergiler dahil − tevkifat ≠ ödenecek tutar (yuvarlama veya ek kalem olabilir)")
    if f.tip in {"TEVKIFAT", "YTBTEVKIFAT"} and not f.tevkifat:
        f.uyarilar.append("Fatura tipi TEVKİFAT ama tevkifat tutarı yok")
    for etiket, no in (("Satıcı", f.satici_vkn), ("Alıcı", f.alici_vkn)):
        if no:
            s = kimlik.dogrula(no, "VKN/TCKN")
            if not s.gecerli:
                f.uyarilar.append(f"{etiket} VKN/TCKN geçersiz: {s.aciklama}")


def dosyalari_topla(girdi: Path):
    """(kaynak adı, bayt) üretir; ZIP'lerin içindeki XML'leri de açar."""
    yollar = [girdi] if girdi.is_file() else sorted(p for p in girdi.rglob("*") if p.is_file())
    for p in yollar:
        uzanti = p.suffix.lower()
        if uzanti == ".xml":
            yield p.name, p.read_bytes()
        elif uzanti == ".zip":
            with zipfile.ZipFile(p) as z:
                for ad in sorted(z.namelist()):
                    if ad.lower().endswith(".xml"):
                        yield f"{p.name}/{ad}", z.read(ad)
                    elif ad.lower().endswith(".zip"):
                        with zipfile.ZipFile(io.BytesIO(z.read(ad))) as ic:
                            for ad2 in sorted(ic.namelist()):
                                if ad2.lower().endswith(".xml"):
                                    yield f"{p.name}/{ad}/{ad2}", ic.read(ad2)


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
UYARI = PatternFill("solid", fgColor="FFF4CE")
PARA = "#,##0.00"


def _sayfa(wb, ad, basliklar, satirlar, para_baslangic=None, genislik=14):
    ws = wb.create_sheet(ad)
    ws.append(basliklar)
    for h in ws[1]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
    for s in satirlar:
        ws.append([float(x) if isinstance(x, Decimal) else x for x in s])
    for i, b in enumerate(basliklar, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(genislik, min(40, len(str(b)) + 2))
        if para_baslangic and i >= para_baslangic:
            for hucre in ws[get_column_letter(i)][1:]:
                if isinstance(hucre.value, float):
                    hucre.number_format = PARA
    ws.freeze_panes = "A2"
    if satirlar:
        ws.auto_filter.ref = ws.dimensions
    return ws


def rapor_yaz(faturalar: list[Fatura], cikti: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    oranlar = sorted({o for f in faturalar for o in f.kdv})
    diger_kodlar = sorted({k for f in faturalar for k in f.diger_vergiler})

    basliklar = ["Kaynak Dosya", "Fatura No", "ETTN", "Tarih", "Senaryo", "Tip", "Para Birimi", "Kur",
                 "Satıcı VKN/TCKN", "Satıcı", "Alıcı VKN/TCKN", "Alıcı", "Mal/Hizmet Toplamı", "İskonto"]
    basliklar += [f"KDV Matrahı %{o:g}" for o in oranlar] + [f"KDV %{o:g}" for o in oranlar]
    basliklar += ["Toplam KDV", "KDV Tevkifatı", "Tevkifat Kodu"]
    basliklar += [VERGI_ADLARI.get(k, f"Vergi Kodu {k}") for k in diger_kodlar]
    basliklar += ["Vergiler Dahil", "Ödenecek", "Satır Sayısı", "Kontrol Uyarıları"]

    satirlar = []
    for f in faturalar:
        s = [f.kaynak, f.no, f.ettn, f.tarih, f.senaryo, f.tip, f.para_birimi, f.kur,
             f.satici_vkn, f.satici_unvan, f.alici_vkn, f.alici_unvan, f.mal_hizmet, f.iskonto]
        s += [f.kdv.get(o, [Decimal(0)])[0] for o in oranlar] + [f.kdv.get(o, [0, Decimal(0)])[1] for o in oranlar]
        s += [f.toplam_kdv, f.tevkifat, ", ".join(f.tevkifat_kodlari)]
        s += [f.diger_vergiler.get(k, Decimal(0)) for k in diger_kodlar]
        s += [f.vergi_dahil, f.odenecek, len(f.satirlar), "; ".join(f.uyarilar)]
        satirlar.append(s)
    ws = _sayfa(wb, "Faturalar", basliklar, satirlar, para_baslangic=13)
    for r, f in enumerate(faturalar, start=2):
        if f.uyarilar:
            for h in ws[r]:
                h.fill = UYARI

    _sayfa(wb, "Satırlar",
           ["Fatura No", "Tarih", "Satıcı", "Alıcı", "Sıra", "Mal/Hizmet", "Miktar", "Birim", "Birim Fiyat",
            "İskonto", "Tutar", "KDV Oranı", "KDV Tutarı", "Para Birimi"],
           [[f.no, f.tarih, f.satici_unvan, f.alici_unvan, s.sira, s.urun, s.miktar, s.birim, s.birim_fiyat,
             s.iskonto, s.tutar, (float(s.kdv_orani) if s.kdv_orani is not None else None), s.kdv_tutari, f.para_birimi]
            for f in faturalar for s in f.satirlar], para_baslangic=9)

    # KDV özeti: para birimi ve oran bazında (TL karşılığıyla); mükerrer ETTN'ler bir kez sayılır
    ozet = defaultdict(lambda: [Decimal(0)] * 4)
    gorulen = set()
    for f in faturalar:
        if f.ettn in gorulen:
            continue
        gorulen.add(f.ettn or id(f))
        kur = f.kur if f.para_birimi != "TRY" else Decimal(1)
        for o, (m, k) in f.kdv.items():
            x = ozet[(f.para_birimi, o)]
            x[0] += m
            x[1] += k
            x[2] += m * kur
            x[3] += k * kur
    _sayfa(wb, "KDV Özeti", ["Para Birimi", "KDV Oranı", "Matrah", "KDV", "Matrah (TL)", "KDV (TL)"],
           [[pb, float(o), *[v.quantize(Decimal("0.01")) for v in vals]] for (pb, o), vals in sorted(ozet.items())],
           para_baslangic=3)

    ettn_sayac = Counter(f.ettn for f in faturalar if f.ettn)
    no_sayac = Counter((f.satici_vkn, f.no) for f in faturalar if f.no)
    kontrol = []
    for f in faturalar:
        for u in f.uyarilar:
            kontrol.append([f.kaynak, f.no, u])
        if f.ettn and ettn_sayac[f.ettn] > 1:
            kontrol.append([f.kaynak, f.no, f"Mükerrer ETTN ({ettn_sayac[f.ettn]} kez)"])
        elif f.no and no_sayac[(f.satici_vkn, f.no)] > 1:
            kontrol.append([f.kaynak, f.no, "Aynı satıcıdan aynı fatura numarası birden çok kez"])
    _sayfa(wb, "Kontroller", ["Kaynak Dosya", "Fatura No", "Uyarı"], kontrol, genislik=30)

    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def calistir(girdi: Path, cikti: Path) -> list[Fatura]:
    if not girdi.exists():
        raise SystemExit(f"Girdi bulunamadı: {girdi}")
    faturalar = [fatura_oku(veri, ad) for ad, veri in dosyalari_topla(girdi)]
    if not faturalar:
        raise SystemExit(f"{girdi} içinde XML veya ZIP dosyası bulunamadı.")
    faturalar.sort(key=lambda f: (f.tarih, f.no))
    rapor_yaz(faturalar, cikti)
    return faturalar


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser(description="UBL-TR e-Fatura / e-Arşiv XML'lerini Excel'e çevirir.")
    p.add_argument("--girdi", type=Path, default=BURASI / "ornek_veri" / "faturalar", help="XML/ZIP klasörü veya tek dosya")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "faturalar.xlsx")
    a = p.parse_args(argv)
    faturalar = calistir(a.girdi, a.cikti)
    uyarili = sum(1 for f in faturalar if f.uyarilar)
    print(f"[OK] {len(faturalar)} belge okundu, {sum(len(f.satirlar) for f in faturalar)} satır")
    print(f"[{'!' if uyarili else 'OK'}] {uyarili} belgede kontrol uyarısı var")
    print(f"[OK] Rapor: {a.cikti.resolve()}")


if __name__ == "__main__":
    sys.exit(main())
