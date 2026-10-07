"""
Kurgusal UBL-TR e-Fatura / e-Arşiv örnekleri üretir (tüm firmalar, kişiler ve numaralar hayalidir).
Yapı, GİB UBL-TR örnek faturalarıyla aynı eleman yollarını kullanır.

Kullanım: python scripts/ornek_efatura_uret.py <hedef_klasör>
"""
import sys
import uuid
import zipfile
from decimal import Decimal
from pathlib import Path

NS_BASLIK = ('xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2" '
             'xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2" '
             'xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"')

# Algoritmaya göre geçerli, kurgusal numaralar
SATICI = ("5260181599", "Örnek Gıda Sanayi ve Ticaret A.Ş.")
ALICI = ("0830166136", "Deneme Lojistik Ltd. Şti.")
ALICI2 = ("6281948215", "Kurgu Tekstil A.Ş.")
KISI_TCKN = "17291716060"


def taraf(etiket, vkn, ad, kisi=False):
    scheme = "TCKN" if len(vkn) == 11 else "VKN"
    if kisi:
        ad_soyad = ad.split(" ", 1)
        isim = f"<cac:Person><cbc:FirstName>{ad_soyad[0]}</cbc:FirstName><cbc:FamilyName>{ad_soyad[1]}</cbc:FamilyName></cac:Person>"
    else:
        isim = f"<cac:PartyName><cbc:Name>{ad}</cbc:Name></cac:PartyName>"
    return (f"<cac:{etiket}><cac:Party><cac:PartyIdentification><cbc:ID schemeID=\"{scheme}\">{vkn}</cbc:ID>"
            f"</cac:PartyIdentification>{isim}<cac:PostalAddress><cbc:CityName>İstanbul</cbc:CityName>"
            f"<cac:Country><cbc:Name>Türkiye</cbc:Name></cac:Country></cac:PostalAddress></cac:Party></cac:{etiket}>")


def kdv_subtotal(matrah, oran, tutar, pb, muafiyet=None):
    neden = f"<cbc:TaxExemptionReasonCode>{muafiyet[0]}</cbc:TaxExemptionReasonCode><cbc:TaxExemptionReason>{muafiyet[1]}</cbc:TaxExemptionReason>" if muafiyet else ""
    return (f"<cac:TaxSubtotal><cbc:TaxableAmount currencyID=\"{pb}\">{matrah}</cbc:TaxableAmount>"
            f"<cbc:TaxAmount currencyID=\"{pb}\">{tutar}</cbc:TaxAmount><cbc:Percent>{oran}</cbc:Percent>"
            f"<cac:TaxCategory>{neden}<cac:TaxScheme><cbc:Name>KDV</cbc:Name><cbc:TaxTypeCode>0015</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal>")


def fatura(no, tarih, senaryo, tip, satici, alici, satirlar, pb="TRY", kur=None, tevkifat=None,
           kisi=False, kdv_bozuk=False, ettn=None, muafiyet=None):
    """satirlar: [(ad, miktar, birim, birim_fiyat, kdv_orani)]"""
    D = Decimal
    satir_xml, oran_toplam = [], {}
    for i, (ad, miktar, birim, fiyat, oran) in enumerate(satirlar, 1):
        tutar = (D(miktar) * D(fiyat)).quantize(D("0.01"))
        kdv = (tutar * D(oran) / 100).quantize(D("0.01"))
        m = oran_toplam.setdefault(oran, [D(0), D(0)])
        m[0] += tutar
        m[1] += kdv
        satir_xml.append(
            f"<cac:InvoiceLine><cbc:ID>{i}</cbc:ID><cbc:InvoicedQuantity unitCode=\"{birim}\">{miktar}</cbc:InvoicedQuantity>"
            f"<cbc:LineExtensionAmount currencyID=\"{pb}\">{tutar}</cbc:LineExtensionAmount>"
            f"<cac:TaxTotal><cbc:TaxAmount currencyID=\"{pb}\">{kdv}</cbc:TaxAmount>{kdv_subtotal(tutar, oran, kdv, pb, muafiyet)}</cac:TaxTotal>"
            f"<cac:Item><cbc:Name>{ad}</cbc:Name></cac:Item><cac:Price><cbc:PriceAmount currencyID=\"{pb}\">{fiyat}</cbc:PriceAmount></cac:Price></cac:InvoiceLine>")
    matrah_top = sum(v[0] for v in oran_toplam.values())
    kdv_top = sum(v[1] for v in oran_toplam.values())
    if kdv_bozuk:  # faturada KDV yanlış yazılmış (kontrolün yakalaması beklenir)
        oran_toplam = {o: [v[0], v[1] + D("15.00")] for o, v in oran_toplam.items()}
        kdv_top += D("15.00")
    alt = "".join(kdv_subtotal(v[0], o, v[1], pb, muafiyet) for o, v in oran_toplam.items())
    tev_xml, tev_tutar = "", D(0)
    if tevkifat:
        kod, yuzde = tevkifat
        tev_tutar = (kdv_top * D(yuzde) / 100).quantize(D("0.01"))
        tev_xml = (f"<cac:WithholdingTaxTotal><cbc:TaxAmount currencyID=\"{pb}\">{tev_tutar}</cbc:TaxAmount>"
                   f"<cac:TaxSubtotal><cbc:TaxAmount currencyID=\"{pb}\">{tev_tutar}</cbc:TaxAmount><cbc:Percent>{yuzde}</cbc:Percent>"
                   f"<cac:TaxCategory><cac:TaxScheme><cbc:TaxTypeCode>{kod}</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal></cac:WithholdingTaxTotal>")
    kur_xml = (f"<cac:PricingExchangeRate><cbc:SourceCurrencyCode>{pb}</cbc:SourceCurrencyCode>"
               f"<cbc:TargetCurrencyCode>TRY</cbc:TargetCurrencyCode><cbc:CalculationRate>{kur}</cbc:CalculationRate></cac:PricingExchangeRate>") if kur else ""
    return (f'<?xml version="1.0" encoding="UTF-8"?>\n<Invoice {NS_BASLIK}>'
            f"<cbc:UBLVersionID>2.1</cbc:UBLVersionID><cbc:CustomizationID>TR1.2</cbc:CustomizationID>"
            f"<cbc:ProfileID>{senaryo}</cbc:ProfileID><cbc:ID>{no}</cbc:ID><cbc:CopyIndicator>false</cbc:CopyIndicator>"
            f"<cbc:UUID>{ettn or str(uuid.uuid5(uuid.NAMESPACE_URL, no)).upper()}</cbc:UUID><cbc:IssueDate>{tarih}</cbc:IssueDate>"
            f"<cbc:InvoiceTypeCode>{tip}</cbc:InvoiceTypeCode><cbc:DocumentCurrencyCode>{pb}</cbc:DocumentCurrencyCode>"
            f"<cbc:LineCountNumeric>{len(satirlar)}</cbc:LineCountNumeric>"
            f"{taraf('AccountingSupplierParty', *satici)}{taraf('AccountingCustomerParty', *alici, kisi=kisi)}{kur_xml}"
            f"<cac:TaxTotal><cbc:TaxAmount currencyID=\"{pb}\">{kdv_top}</cbc:TaxAmount>{alt}</cac:TaxTotal>{tev_xml}"
            f"<cac:LegalMonetaryTotal><cbc:LineExtensionAmount currencyID=\"{pb}\">{matrah_top}</cbc:LineExtensionAmount>"
            f"<cbc:TaxExclusiveAmount currencyID=\"{pb}\">{matrah_top}</cbc:TaxExclusiveAmount>"
            f"<cbc:TaxInclusiveAmount currencyID=\"{pb}\">{matrah_top + kdv_top}</cbc:TaxInclusiveAmount>"
            f"<cbc:PayableAmount currencyID=\"{pb}\">{matrah_top + kdv_top - tev_tutar}</cbc:PayableAmount></cac:LegalMonetaryTotal>"
            f"{''.join(satir_xml)}</Invoice>")


def main(hedef: Path) -> None:
    hedef.mkdir(parents=True, exist_ok=True)
    f1 = fatura("ORN2026000000101", "2026-09-02", "TEMELFATURA", "SATIS", SATICI, ALICI, [
        ("Ayçiçek Yağı 5 lt", 120, "C62", "289.90", 1),
        ("Bulaşık Deterjanı 1 lt", 48, "C62", "74.50", 20),
        ("Nakliye Hizmeti", 1, "C62", "1500.00", 20),
    ])
    (hedef / "ORN2026000000101.xml").write_text(f1, encoding="utf-8")
    (hedef / "ORN2026000000102.xml").write_text(fatura(
        "ORN2026000000102", "2026-09-05", "TICARIFATURA", "TEVKIFAT", SATICI, ALICI2,
        [("Temizlik Hizmeti (Eylül)", 1, "MON", "20000.00", 20)], tevkifat=("624", 20)), encoding="utf-8")
    (hedef / "EAR2026000000031.xml").write_text(fatura(
        "EAR2026000000031", "2026-09-07", "EARSIVFATURA", "SATIS", SATICI, (KISI_TCKN, "Ayşe Yılmaz"),
        [("Un 25 kg", 2, "C62", "640.00", 1)], kisi=True), encoding="utf-8")
    (hedef / "IHR2026000000007.xml").write_text(fatura(
        "IHR2026000000007", "2026-09-10", "IHRACAT", "ISTISNA", SATICI, ("9960308245", "Example Trading GmbH"),
        [("Kuru Kayısı 10 kg Koli", 200, "C62", "38.50", 0)], pb="USD", kur="33.8750",
        muafiyet=("301", "11/1-a Mal İhracatı")), encoding="utf-8")
    (hedef / "ORN2026000000110.xml").write_text(fatura(
        "ORN2026000000110", "2026-09-12", "TEMELFATURA", "SATIS", SATICI, ("1234567890", "Hayali Kırtasiye"),
        [("A4 Fotokopi Kağıdı", 40, "PA", "165.00", 20)], kdv_bozuk=True), encoding="utf-8")
    with zipfile.ZipFile(hedef / "gib_portal_indirme.zip", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("ORN2026000000115.xml", fatura(
            "ORN2026000000115", "2026-09-15", "TEMELFATURA", "SATIS", SATICI, ALICI,
            [("Pirinç 1 kg", 300, "KGM", "62.00", 1)]))
        z.writestr("ORN2026000000101_kopya.xml", f1)   # aynı ETTN: mükerrer
    (hedef / "bozuk.xml").write_text("<Invoice><cbc:ID>yarim", encoding="utf-8")
    print(f"[OK] örnek faturalar -> {hedef}")


if __name__ == "__main__":
    main(Path(sys.argv[1]))
