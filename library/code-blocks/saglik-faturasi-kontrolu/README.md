# Sağlık Faturası Kontrolü · Kod Bloğu

> Sigortacılık › Sağlık Sigortaları › Provizyon Uzmanı · Workers / Workless

Anlaşmalı kurum faturalarını satır bazında karşılaştırır: anlaşmalı fiyat listesi, paket içerikleri ve onaylı
provizyon. Kesinti tutarlarını, kabul edilen tutarı ve şirket / sigortalı payını hesaplar. Kuruma gönderilecek
kesinti listesini hazırlar. İnternete bağlanmaz.

## Ne kontrol eder?

| Kontrol | Sonuç |
|---|---|
| Birim fiyat anlaşmalı fiyattan yüksek | (fark × adet) kesinti |
| Tutar, adet × birim fiyattan büyük (hesap hatası) | Fark kesinti |
| Paket fiyatına dahil hizmet aynı provizyonda ayrıca faturalanmış (ör. ameliyat paketi + ameliyathane + yatak) | Satırın tamamı kesinti |
| Mükerrer satır: aynı provizyon, kod, tarih, adet ve fiyat | Tekrarlayan satır kesinti |
| Adet sınırı: fiyat listesindeki "Maks Adet" provizyon başına aşılmış | Aşan adet kesinti |
| Hizmet fiyat listesinde yok (ilaç, malzeme vb.) | İnceleme |
| İşlem tarihi provizyon tarihinden önce | İnceleme |
| Provizyon reddedilmiş | Fatura kabul edilmez |
| Provizyon incelemede veya bulunamadı | Beklemede |

**Fatura düzeyi:**
- **Şirket payı:** kabul edilen tutar × (1 − katılım payı). Onaylı provizyon tutarıyla sınırlıdır; aşan kısım
  sigortalı payına eklenir ve uyarılır.
- **Kesinti oranı:** %10 veya üstündeyse kurumla mutabakat uyarısı verilir.

Provizyon dosyası, `provizyon-talebi-degerlendirme` paketinin "Değerlendirme" sayfasıyla aynı sütunları kabul
eder: Talep No, Talep Tarihi, Ön Karar, Ödenecek, Katılım Payı %.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 6 fatura, 15 satır
python main.py --faturalar fatura_kalemleri.xlsx --fiyatlar fiyat_listesi.xlsx --provizyonlar provizyonlar.xlsx --paketler paket_icerikleri.csv
```

| Dosya | Sütunlar |
|---|---|
| Fatura kalemleri | Fatura No, Kurum, Provizyon No, İşlem Tarihi, Hizmet Kodu, Hizmet Adı, Adet, Birim Fiyat (TL), Tutar (TL) |
| Fiyat listesi | Kurum, Hizmet Kodu, Hizmet Adı, Anlaşmalı Fiyat (TL), Maks Adet |
| Provizyonlar | Talep No, Talep Tarihi, Sigortalı No, Ön Karar, Ödenecek, Katılım Payı % |
| Paket içerikleri (isteğe bağlı) | Paket Kodu, Dahil Hizmet Kodu |

Fiyat listesinde kurum boş bırakılırsa satır tüm kurumlar için geçerlidir.

## Çıktı

`saglik_fatura_kontrolu.xlsx`:
- `Fatura Özeti`: fatura bazında şunlar ve boş "Uzman Onayı" sütunu; altta toplamlar:
  - fatura tutarı, kesinti, beklemede / inceleme tutarı, kabul edilen tutar
  - katılım payı, şirket payı, sigortalı payı
  - onaylı provizyon tutarı ve durum
- `Satır Kontrolü`: her satırın anlaşmalı fiyatı, kesintisi, kabul tutarı, durumu ve nedenleri.
- `Kesinti Listesi`: kuruma gönderilecek kesintiler ve gerekçeleri; boş "Kurum İtirazı" sütunu.
- `Uyarılar`.

## Dikkat

- **Kaynaklar:** Fiyat listesi, paket içerikleri ve adet sınırları kurumla yapılan anlaşmadan alınmalıdır. Örnek
  değerler kurgusaldır.
- **İlaç ve malzeme:** İlaç ve tıbbi malzeme kalemleri fiyat listesinde yoksa "İnceleme"ye düşer. Bu kalemler
  anlaşmanın ilaç / malzeme fiyatlandırma kuralına göre ayrıca kontrol edilmelidir (ör. fatura, barkod, kâr oranı).
- **Tıbbi uygunluk:** Hizmetin tıbbi gerekliliği ve tanıyla uyumu bu aracın kapsamında değildir. Provizyon
  uzmanı / şirket hekimi değerlendirir.
- **Örnek veri:** Örnek kurum, faturalar ve fiyatlar kurgusaldır.

## Testler

Şunlar test edilir:
- Her kontrol türü ve kesinti tutarları.
- Provizyon reddi, inceleme ve bulunamama durumları.
- Katılım payı ve provizyon tavanı; fatura toplamları.
- Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
