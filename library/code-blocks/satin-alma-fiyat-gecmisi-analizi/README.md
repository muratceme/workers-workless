# Satın Alma Fiyat Geçmişi Analizi · Kod Bloğu

> Satın Alma › Satın Alma Uzmanı · Workers / Workless

Geçmiş alımlarınızı malzeme bazında sıraya dizer ve her fiyat artışını **aynı dönemdeki enflasyon ve kur
hareketiyle** karşılaştırır. "Bu zam enflasyon kadar mı, yoksa üstünde mi?" sorusunu sayıyla yanıtlar; endeksin
belirgin üstündeki artışları işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| TL birim fiyat | Birim fiyat × alım ayının kuru (TL alımlarda kur 1) |
| Önceki alıma göre % | TL fiyat ÷ aynı malzemenin önceki alımındaki TL fiyat − 1 |
| Endeks değişimi % | Alım ayının endeksi ÷ önceki alım ayının endeksi − 1 |
| Fark (puan) | Fiyat değişimi − endeks değişimi |
| Olağandışı artış | Fark `--esik` (10) puandan büyükse; endeks yoksa fiyat değişimi eşikten büyükse |
| Nominal ve reel değişim | Son fiyat ÷ ilk fiyat − 1; reel = (1 + nominal) ÷ (1 + endeks değişimi) − 1 |
| Ağırlıklı ortalama fiyat | Σ(miktar × TL fiyat) ÷ Σ miktar |
| Döviz fiyatı değişimi | Dövizli alımlarda fiyatın kendi para birimindeki değişimi; TL artışının ne kadarının kurdan geldiğini gösterir |

**Örnek:** Koli 4,45 TL'den 5,40 TL'ye çıkmış: %21,35 artış. Aynı dönemde endeks 111'den 116'ya çıkmış: %4,50.
Fark 16,84 puan → **olağandışı artış**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                    # örnek veri
python main.py --alimlar alimlar.xlsx
python main.py --alimlar alimlar.xlsx --endeks endeks.xlsx --kurlar kurlar.xlsx --esik 5
```

Alım dosyasında zorunlu sütunlar: Tarih, Malzeme, Birim Fiyat. Diğerleri: Malzeme Kodu, Tedarikçi, Miktar, Para
Birimi (boşsa TL). Endeks dosyası: Dönem (AA.YYYY), Endeks. Kur dosyası: Dönem, Para Birimi, Kur. Dövizli alımın
ayına ait kur yoksa o alım analize alınmaz ve uyarılır.

## Çıktı

`cikti/fiyat_gecmisi_analizi.xlsx`:

- `Malzeme Özeti`: alım sayısı, toplam miktar ve tutar, ağırlıklı ortalama, ilk ve son fiyat, en düşük ve en yüksek,
  nominal, endeks ve reel değişim, tedarikçi sayısı. Olağandışı artışı olan malzemeler en üsttedir.
- `Alım Geçmişi`: her alımın TL fiyatı, önceki alıma göre değişimi, endeks değişimi, farkı ve işareti.
- `Uyarılar`: olağandışı artışlar (tedarikçi değişikliği notuyla), eksik kur ve endeks, tek alımlı malzemeler.

## Dikkat

- **Endeksi ve kurları siz seçersiniz.** Kod veri çekmez. Hangi endeksin (genel üretici fiyatları, sektör
  endeksi, emtia fiyatı) malzemeniz için anlamlı olduğuna siz karar verirsiniz; genel endeks, tek bir hammaddenin
  fiyat hareketini açıklamayabilir. Örnek endeks ve kurlar kurgusaldır.
- **İşaret bir inceleme önerisidir.** Miktar iskontosu, vade, teslim şekli, kalite veya ambalaj değişikliği fiyatı
  meşru biçimde değiştirir; kod bunları bilmez.
- **Karşılaştırma aylıktır.** Endeks ve kur ay bazında eşleştirilir; ay içi hareketler ve fatura ile sipariş tarihi
  arasındaki fark dikkate alınmaz.
- **Aynı malzeme, aynı birim.** Malzemeler koda (kod yoksa ada) göre gruplanır; birim değişikliği (adet → koli)
  ayırt edilmez. Farklı birimli alımları ayrı kodla girin.
- **Örnek veri:** Malzemeler, tedarikçiler, fiyatlar, endeks ve kurlar kurgusaldır.

## Testler

- Fiyat, endeks ve fark hesapları; nominal, reel ve ağırlıklı ortalama (elle hesaplanmış); eşik sınırı.
- Dövizli alım, kur etkisi, eksik kur ve endeks, tedarikçi değişimi, tek alımlı malzeme.
- Dönem ve sayı okuma, okuma hataları, komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
