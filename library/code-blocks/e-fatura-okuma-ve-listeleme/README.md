# e-Fatura Okuma ve Listeleme · Kod Bloğu

> Muhasebe › Ön Muhasebe Elemanı · Mali Müşavirlik Bürosu › Muhasebe Elemanı · Workers / Workless

GİB **UBL-TR** biçimindeki e-Fatura ve e-Arşiv XML dosyalarını okur ve tek bir Excel dosyasında
listeler: fatura başlıkları, **KDV oranı kırılımı**, **tevkifat**, diğer vergiler (ÖTV, ÖİV vb.),
satır detayları ve tutarlılık kontrolleri. GİB portalından veya özel entegratörden indirilen
**ZIP dosyalarını** doğrudan açar. İnternete bağlanmaz.

## Ne üretir? (`cikti/faturalar.xlsx`)

| Sayfa | İçerik |
|---|---|
| Faturalar | Fatura no, ETTN, tarih, senaryo (TEMELFATURA, TICARIFATURA, EARSIVFATURA, IHRACAT…), tip (SATIS, IADE, TEVKIFAT, ISTISNA…), para birimi ve kur, satıcı/alıcı VKN-TCKN ve unvan, oran bazında KDV matrahı ve KDV, tevkifat tutarı ve kodu, diğer vergiler, vergiler dahil ve ödenecek tutar |
| Satırlar | Her fatura satırı: mal/hizmet, miktar, birim (C62 → Adet, KGM → kg…), birim fiyat, iskonto, tutar, KDV oranı ve tutarı |
| KDV Özeti | Para birimi ve KDV oranı bazında matrah ve KDV toplamı; dövizli faturalar için TL karşılığı. Mükerrer faturalar bir kez sayılır |
| Kontroller | Okunamayan XML, mükerrer ETTN, aynı satıcıdan tekrar eden fatura no, matrah × oran ≠ KDV, vergiler dahil − tevkifat ≠ ödenecek, satır toplamı farkı, geçersiz VKN/TCKN |

Uyarı içeren faturalar **Faturalar** sayfasında sarı görünür.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                     # örnek faturalarla dener
python main.py --girdi "C:/Faturalar/2026-09"       # klasör (alt klasörler ve ZIP'ler dahil)
python main.py --girdi gelen_faturalar.zip --cikti cikti/eylul.xlsx
```

## Notlar ve sınırlamalar

- Yalnızca UBL-TR **XML** okunur; PDF faturalar için bu görevin AI Agent sürümüne bakın.
- İrsaliye (DespatchAdvice) ve uygulama yanıtı belgeleri atlanır ve "Kontroller"de belirtilir.
- Kontroller bilgilendirme amaçlıdır; GİB şematron doğrulamasının yerini tutmaz.
- Vergi türü adları GİB vergi kodu tablosundan alınmıştır; adı doğrulanmamış kodlar "Vergi Kodu XXXX" olarak görünür.

## Testler

```bash
python -m unittest discover -s tests
```

Testler kurgusal örneklerin yanında GİB'in resmî tevkifatlı örnek faturasını da (`tests/fixtures/`) kullanır.

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
