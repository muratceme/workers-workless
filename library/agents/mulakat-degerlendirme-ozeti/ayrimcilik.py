"""
ayrimcilik.py — Workers / Workless ortak ayrımcı ifade tarayıcısı (iş ilanı, mülakat soruları)

Kaynağı: library/_ortak/ayrimcilik.py; paket klasörlerindeki kopyalar CI tarafından bununla aynı tutulur.

Dayanak (genel): 4857 sayılı İş Kanunu md. 5 (eşit davranma), 6701 sayılı Türkiye İnsan Hakları ve Eşitlik
Kurumu Kanunu md. 3 ve 6 (cinsiyet, ırk, renk, dil, din, inanç, mezhep, felsefi ve siyasi düşünce, etnik köken,
servet, doğum, medeni hâl, sağlık durumu, engellilik ve yaş temelinde istihdamda ayrımcılık yasağı; işin
niteliği gereği zorunlu farklı muamele md. 7 kapsamında değerlendirilebilir), 6356 sayılı Kanun md. 25
(sendika üyeliği nedeniyle ayrım yasağı), KVKK md. 4 (ölçülülük — gereksiz kişisel veri istenmemesi).

Bu tarayıcı kural tabanlı bir uyarı aracıdır; hukuki değerlendirme yerine geçmez. "yüksek" bulgular doğrudan
ayrımcılık riski taşır; "dikkat" bulguları işin gereği olarak gerekçelendirilebiliyorsa kullanılabilir.
"""
from __future__ import annotations

import re

# (seviye, kategori, desen, açıklama)
KURALLAR = [
    ("yüksek", "Yaş", r"\b\d{2}\s*(?:-|–|ile)\s*\d{2}\s*yaş", "Yaş aralığı şartı"),
    ("yüksek", "Yaş", r"\b\d{2}\s*yaşın(?:ı|ın)?\s*(?:altında|üstünde|altı|üstü|geçmemiş|aşmamış|doldurmamış)", "Yaş sınırı"),
    ("yüksek", "Yaş", r"\byaş sınır", "Yaş sınırı"),
    ("yüksek", "Yaş", r"\bgenç (?:ve |, ?)?(?:dinamik|eleman|personel|ekip arkadaşı|aday)", "Yaşa dayalı ifade ('genç')"),
    ("yüksek", "Cinsiyet", r"\bbayan\b", "Cinsiyete dayalı ifade ('bayan')"),
    ("yüksek", "Cinsiyet", r"\b(?:sadece |yalnızca |yalnız )?(?:erkek|kadın|bayan) (?:eleman|personel|çalışan|aday|adaylar|garson|şoför|sekreter|operatör)",
     "Cinsiyet şartı"),
    ("yüksek", "Cinsiyet", r"\b(?:erkek|kadın) olmak\b", "Cinsiyet şartı"),
    ("yüksek", "Medeni hâl ve aile", r"\b(?:evli|bekar|bekâr)\b", "Medeni hâl şartı/sorusu"),
    ("yüksek", "Medeni hâl ve aile", r"medeni (?:hal|hâl|durum)", "Medeni hâl"),
    ("yüksek", "Medeni hâl ve aile", r"\bçocu(?:ğu|k|ğunuz|klarınız)\b.{0,25}\b(?:var mı|yok|olmayan|planı|düşün)", "Çocuk / aile planı"),
    ("yüksek", "Medeni hâl ve aile", r"\b(?:hamile|gebe|hamilelik|gebelik)", "Gebelik"),
    ("yüksek", "Görünüş", r"(?:dış görünüş|güzel (?:görünümlü|yüzlü)|yakışıklı|hoş görünümlü|alımlı|fiziği düzgün|"
                         r"boy\w*\s*(?:en az\s*)?\d{3}|\bkilo(?:lu|su|nuz)\b)", "Fiziksel görünüş şartı"),
    ("yüksek", "Sağlık ve engellilik", r"(?:sağlık sorunu olmayan|herhangi bir sağlık sorunu|engel(?:li|i) olmayan|engeli bulunmayan|"
                                       r"kronik rahatsızlığı olmayan|sağlıklı (?:olmak|olan|aday|birey|bünye)|rahatsızlığınız var mı|"
                                       r"hastalığınız var mı|ilaç kullanıyor)", "Sağlık durumu / engellilik şartı"),
    ("yüksek", "Din ve inanç", r"(?:başörtü|türban|tesettür|\bnamaz|\boruç|dini (?:inanç|görüş|tercih|vecibe)|dininiz|inancınız|"
                               r"\bmezhe|\balevi|\bsünni|ibadet)", "Din / inanç / kıyafet"),
    ("yüksek", "Etnik köken ve dil", r"(?:\bırk(?:ı|ınız)?\b|etnik köken|\bkürt|türk asıllı|nerelisiniz|memleketiniz|hemşehri|aslen nereli)",
     "Etnik köken / köken sorusu"),
    ("dikkat", "Etnik köken ve dil", r"(?:anadili|ana dili)", "Anadili şartı (işin gereği değilse dil yeterliliği yazın)"),
    ("yüksek", "Siyasi görüş ve sendika", r"(?:siyasi görüş|hangi partiy|parti üyeliği|sendika üye|sendikalı|sendikaya)", "Siyasi görüş / sendika üyeliği"),
    ("dikkat", "Askerlik", r"askerli(?:ğini|k)\s*(?:yapmış|tamamlamış|ile ilişiği|ilişiği|durumu|tecilli|muaf)", "Askerlik şartı (dolaylı cinsiyet ayrımı riski)"),
    ("dikkat", "Uyruk", r"(?:t\.?c\.? vatandaşı|türk vatandaşı|uyruklu|vatandaşı olmak)", "Uyruk şartı (yasal zorunluluk yoksa)"),
    ("dikkat", "Kişisel veri", r"(?:fotoğraflı|fotoğraf ekli|fotoğrafınızı|tc kimlik|t\.c\. kimlik|nüfus cüzdan|kan grubu)",
     "Başvuruda gereksiz kişisel veri (KVKK ölçülülük)"),
    ("dikkat", "Kişisel veri", r"(?:sabıka kaydı|adli sicil)", "Adli sicil kaydı (yalnız mevzuat/işin gereği varsa)"),
    ("dikkat", "Yaşam tarzı", r"sigara (?:içmeyen|kullanmayan)", "Yaşam tarzı şartı"),
    ("dikkat", "Ehliyet", r"\b(?:b sınıfı )?ehliyet", "Ehliyet şartı (iş gerektirmiyorsa dolaylı ayrım)"),
]
_DESENLER = [(s, k, re.compile(d, re.I), a) for s, k, d, a in KURALLAR]


def kucuk(s: str) -> str:
    return str(s or "").replace("İ", "i").replace("I", "ı").lower()


def tara(metin: str) -> list[dict]:
    """Metindeki riskli ifadeleri döndürür: [{seviye, kategori, ifade, aciklama}]"""
    k = kucuk(metin)
    bulgular, gorulen = [], set()
    for seviye, kategori, desen, aciklama in _DESENLER:
        for m in desen.finditer(k):
            anahtar = (kategori, m.group(0).strip())
            if anahtar in gorulen:
                continue
            gorulen.add(anahtar)
            bulgular.append({"seviye": seviye, "kategori": kategori, "ifade": m.group(0).strip(), "aciklama": aciklama})
    return bulgular


def ozet(bulgular: list[dict]) -> str:
    return "; ".join(f"[{b['seviye']}] {b['kategori']}: '{b['ifade']}'" for b in bulgular)
