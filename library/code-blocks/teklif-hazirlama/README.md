# Teklif Hazırlama · Kod Bloğu

> Satış › Satış Temsilcisi · Workers / Workless

Fiyat listesi, iskonto kuralları ve seçilen ürünlerden müşteriye gönderilecek **teklif dosyasını** hazırlar.
Ayrıca şirket içinde kalacak bir **iç kontrol dosyası** yazar: iskonto zinciri, marj ve yönetici onayı gerekip
gerekmediği. İnternete bağlanmaz.

## Nasıl hesaplar?

1. **Döviz:** Liste fiyatı teklif dövizine çevrilir: fiyat × ürün dövizinin kuru / teklif dövizinin kuru. Kurlar
   teklif bilgisi dosyasından okunur (`Kur EUR`, `Kur USD`...). Gereken kur yoksa teklif hazırlanmaz.
2. **İskonto kuralları:** Koşulu ve geçerlilik tarihi uyan kurallar uygulanır.
   - Müşteri grubu, kategori ve ürün kuralları koşula göre uygulanır. Miktar kuralı, "Min Miktar" ve üzeri satırlara
     uygulanır.
   - **Aynı "Grup" içindeki kurallardan yalnız en yüksek oran** geçerlidir. Örneğin vana 20+ %5 ve 50+ %10 kuralları
     aynı gruptadır; 60 adette yalnız %10 uygulanır.
   - Farklı grupların iskontoları **zincirleme** uygulanır: net = liste × (1 − i₁) × (1 − i₂) … Satırdaki ek (manuel)
     iskonto zincirin sonuna eklenir.
   - Örnek: 850 EUR × 48,10 = 40.885 TL; bayi %15 ve kampanya %5 ile 40.885 × 0,85 × 0,95 = 33.014,64 TL.
3. **Tutarlar:** Net birim fiyat kuruşa yuvarlanır; tutar = miktar × net birim fiyat. KDV, oran bazında ara toplam
   üzerinden hesaplanır.
4. **İç kontrol:**
   - Marj = (net − maliyet) / net. Maliyet de teklif dövizine çevrilir.
   - Yönetici onayı şu durumlarda gerekir:
     - marj `--min-marj` (%15) altında;
     - toplam iskonto `--max-iskonto` (%25) üstünde;
     - fiyat listesinde olmayan veya miktarı geçersiz kalem var.

| Uyarı | Önem |
|---|---|
| Fiyat listesinde olmayan kalem; geçersiz miktar (kalem alınmaz) | Yüksek |
| Eksi marj | Yüksek |
| Marj alt sınırın altında; toplam iskonto sınırın üstünde | Orta |
| Asgari sipariş miktarının altında | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek: 8 kalemlik bayi teklifi
python main.py --fiyat-listesi fiyat.xlsx --iskontolar iskonto.xlsx --teklif kalemler.xlsx --bilgi teklif_bilgisi.csv --cikti teklifler
```

| Dosya | Sütunlar |
|---|---|
| Fiyat listesi | Ürün Kodu, Ürün, Kategori, Birim, Liste Fiyatı, Döviz, KDV Oranı, Maliyet (iç, isteğe bağlı), Asgari Sipariş |
| İskonto kuralları | Kural, Tür (Müşteri grubu / Kategori / Ürün / Miktar / Kampanya), Koşul, Min Miktar, İskonto, Grup, Başlangıç, Bitiş |
| Teklif kalemleri | Ürün Kodu, Miktar, Ek İskonto, Not |
| Teklif bilgisi | İki sütun (Alan;Değer): Firma, Teklif No, Tarih, Müşteri, Müşteri Grubu, İlgili Kişi, Geçerlilik (gün), Döviz, Kur EUR, Kur USD, Ödeme, Teslim Şekli, Teslim Süresi, Hazırlayan, Notlar |

## Çıktı

- `teklif_<no>.xlsx` müşteriye gönderilir:
  - firma ve teklif bilgileri, geçerlilik tarihi;
  - satırlar: liste fiyatı, iskonto oranı, net birim fiyat, tutar;
  - liste toplamı, iskonto, ara toplam, KDV dökümü, genel toplam;
  - ödeme ve teslim şartları.

  Dosya sayfaya sığacak biçimde ayarlıdır ve PDF olarak kaydedilebilir.
- `teklif_<no>_ic_kontrol.xlsx` şirket içinde kalır: özet (ortalama iskonto, brüt marj, onay), satır bazında iskonto
  zinciri ve marj, uyarılar.

## Dikkat

- **İç kontrol dosyasını müşteriye göndermeyin.** Maliyet ve marj bilgisi yalnız bu dosyadadır.
- **Kur riski:** Döviz bazlı ürünlerde TL teklif, teklif tarihindeki kurla hesaplanır. Geçerlilik süresince kur
  değişebilir; teklif şartlarına kur farkı maddesi eklemeyi değerlendirin.
- **İskonto politikası** şirketindir. Kural dosyasındaki oranlar örnektir.
- **Örnek veri:** Örnek firma, ürünler, fiyatlar ve kurlar kurgusaldır.

## Testler

Şunlar test edilir:
- Döviz çevrimi, zincirleme iskonto, aynı grupta en yüksek oran, süresi geçmiş kampanya (elle hesaplanmış).
- Ara toplam, KDV, genel toplam; düşük marj ve yönetici onayı; asgari sipariş; listede olmayan ürün.
- Müşteri dosyasında maliyet ve marj bilgisi olmaması.
- Eksik kur hatası, farklı müşteri grubu ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
