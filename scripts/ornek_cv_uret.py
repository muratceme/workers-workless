"""
Örnek (kurgusal) CV dosyalarını üretir: TXT, DOCX ve PDF.
Tüm kişiler, şirketler ve iletişim bilgileri hayalidir.

Kullanım: python scripts/ornek_cv_uret.py <hedef_klasör>
"""
import sys
import zipfile
from pathlib import Path

CVLER = {
    "ayse_yilmaz.txt": """Ayşe Yılmaz
İK Uzmanı · İstanbul
ayse.yilmaz@ornek-mail.com | 0532 111 22 33 | linkedin.com/in/ayse-yilmaz-ornek

DENEYİM
Kıdemli İK Uzmanı — Örnek Holding A.Ş.
03/2019 - Günümüz
İşe alım süreçlerinin uçtan uca yönetimi, mülakat planlama, performans yönetimi.

İK Uzmanı — Deneme Tekstil
2015 - 2019
Bordro, SGK bildirgeleri ve eğitim planlama.

EĞİTİM
Marmara Üniversitesi, İşletme (Lisans)
2010 - 2014

BECERİLER
Excel, Power BI, SAP, İş hukuku, Bordro, İşe alım
YABANCI DİL
İngilizce (C1), Almanca (A2)
""",
    "mehmet_kaya.docx": """MEHMET KAYA
mehmet.kaya@ornek-mail.com
+90 (533) 444 55 66
Deneyim
Finansal Analist, Hayali Bank, 2020 - halen
Bütçe, raporlama, nakit yönetimi. SQL ve Python ile otomatik raporlama.
Muhasebe Uzmanı, Kurgu Lojistik, 2017 - 2020
Mali tablo hazırlama, vergi beyannameleri, Logo ERP.
Eğitim
Boğaziçi Üniversitesi, Ekonomi Yüksek Lisans, 2015 - 2017
Ankara Üniversitesi, İktisat Lisans, 2011 - 2015
Diller
English (Advanced)
""",
    "zeynep_demir.pdf": """Zeynep Demir
zeynep.demir@ornek-mail.com  0544 777 88 99
Experience
Sales Specialist - Ornek Perakende, 2022 - present
Musteri iliskileri, satis, CRM (Salesforce).
Education
Anadolu Universitesi, Isletme Lisans, 2017 - 2021
Languages: English, Spanish
""",
    "bozuk_dosya.pdf": None,  # okunamayan dosya senaryosu
}


def docx_yaz(yol: Path, metin: str) -> None:
    from xml.sax.saxutils import escape

    paragraflar = "".join(f"<w:p><w:r><w:t xml:space=\"preserve\">{escape(s)}</w:t></w:r></w:p>" for s in metin.splitlines())
    belge = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
             f"<w:body>{paragraflar}</w:body></w:document>")
    with zipfile.ZipFile(yol, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", belge)


def pdf_yaz(yol: Path, metin: str) -> None:
    """Tek sayfalık, yalnızca ASCII metin içeren minimal PDF."""
    def kac(s):
        return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    satirlar = metin.splitlines()
    akis = "BT /F1 11 Tf 50 790 Td 14 TL " + " ".join(f"({kac(s)}) '" for s in satirlar) + " ET"
    nesneler = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        f"<< /Length {len(akis)} >>\nstream\n{akis}\nendstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    cikti, ofsetler = b"%PDF-1.4\n", []
    for i, n in enumerate(nesneler, 1):
        ofsetler.append(len(cikti))
        cikti += f"{i} 0 obj\n{n}\nendobj\n".encode("latin-1")
    xref = len(cikti)
    cikti += f"xref\n0 {len(nesneler) + 1}\n0000000000 65535 f \n".encode()
    cikti += "".join(f"{o:010d} 00000 n \n" for o in ofsetler).encode()
    cikti += f"trailer\n<< /Size {len(nesneler) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    yol.write_bytes(cikti)


def main(hedef: Path) -> None:
    hedef.mkdir(parents=True, exist_ok=True)
    for ad, metin in CVLER.items():
        yol = hedef / ad
        if metin is None:
            yol.write_bytes(b"%PDF-1.4\nbu dosya bilerek bozuk\n")
        elif ad.endswith(".docx"):
            docx_yaz(yol, metin)
        elif ad.endswith(".pdf"):
            pdf_yaz(yol, metin)
        else:
            yol.write_text(metin, encoding="utf-8")
    print(f"[OK] {len(CVLER)} örnek CV -> {hedef}")


if __name__ == "__main__":
    main(Path(sys.argv[1]))
