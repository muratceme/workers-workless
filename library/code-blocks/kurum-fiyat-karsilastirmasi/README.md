# Kurum Fiyat Karşılaştırması · Kod Bloğu

> Sağlık Sigortaları › Anlaşmalı Kurumlar Uzmanı · Workers / Workless

Anlaşmalı sağlık kurumlarının işlem fiyatlarını **aynı bölge ve kurum grubundaki emsalleriyle** karşılaştırır; her
kurum için tek bir fiyat endeksi üretir, emsallerinden belirgin pahalı kurum ve işlemleri işaretler. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Emsal kümesi | Aynı bölge + aynı kurum grubu + aynı işlem kodu; en az `--min` (3) kurum varsa karşılaştırılır |
| Ortanca fiyat | Kümedeki fiyatların ortancası (çift sayıda ise ortadaki iki fiyatın ortalaması) |
| İşlem endeksi | Kurumun fiyatı ÷ ortanca × 100 |
| Kurum endeksi | Σ (fiyat × adet) ÷ Σ (ortanca × adet) × 100; yalnız karşılaştırılabilen işlemler |
| Fiyat yayılımı % | (En yüksek − en düşük) ÷ ortanca × 100 |
| Ortanca üstü tutar | Σ max(fiyat − ortanca, 0) × adet |
| Durum | Kurum endeksi > `--esik` (120) "Pahalı"; < 200 − eşik (80) "Ucuz"; arası "Emsal düzeyinde" |
| Aşırı fiyat | İşlem endeksi > `--asiri` (150) |

`Adet` sütunu yoksa her işlem 1 adet sayılır; yani kurum endeksi fiyatların düz toplamına göre hesaplanır.

**Örnek:** Dört kurumun muayene fiyatları 2.000, 2.200, 2.400 ve 3.200 → ortanca **2.300**. 3.200 isteyen kurumun
işlem endeksi **139,13**. Aynı kurumun dört işlemi adetleriyle ağırlıklandırıldığında kurum endeksi **132,36** →
"Pahalı".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                           # örnek veri
python main.py --fiyatlar kurum_fiyatlari.xlsx
python main.py --fiyatlar kurum_fiyatlari.xlsx --esik 115 --asiri 140 --min 5
```

Zorunlu sütunlar: Kurum, İşlem Kodu, Fiyat. Aynı kurum ve işlem iki kez geçerse ikinci satır alınmaz.

## Çıktı

`cikti/kurum_fiyat_karsilastirmasi.xlsx`:

- `Kurumlar`: endekse göre sıralı; durum, karşılaştırılan işlem sayısı, ortanca üstü tutar, en yüksek endeksli işlem.
- `İşlemler`: emsal kümesi bazında en düşük, ortanca, en yüksek fiyat ve yayılım.
- `Fiyatlar`: her fiyat satırının ortancası, farkı ve endeksi.
- `Uyarılar`.

## Dikkat

- **Fiyatların aynı şeyi kapsadığını kod bilemez.** Paket fiyatın içeriği (yatış günü, malzeme, hekim ücreti,
  anestezi), KDV ve katılım payı kurumdan kuruma değişebilir. Karşılaştırmadan önce işlem tanımlarını ve kapsamı
  eşitleyin; aynı kodu taşıyan farklı kapsamlı fiyatlar yanıltıcı endeks üretir.
- **Emsal tanımı sizin sınıflandırmanızdır.** Bölge ve kurum grubunu nasıl yazdıysanız küme odur; ilçe, kurum türü
  ya da hizmet kalitesi farkı ayrıca gözetilmez.
- **Ortanca "doğru fiyat" değildir.** Kümedeki bütün kurumlar pahalıysa ortanca da pahalıdır. Endeks, kurumun
  emsallerine göre yerini gösterir; anlaşma fiyatının makul olup olmadığına karar vermez.
- **"Ortanca üstü tutar" tasarruf tahmini değildir.** Fiyat ortancaya çekilseydi adetler aynı kalırdı varsayımına
  dayanan kaba bir büyüklüktür; pazarlık gücünü, hasta yönlendirmesini ve kalite farkını içermez.
- **Az kurumlu kümede sonuç oynaktır.** Üç kurumlu kümede ortanca ortadaki kurumun fiyatıdır; `--min` ile eşiği
  yükseltebilirsiniz.
- **Kapsam dışı:** resmî fiyat tarifeleriyle karşılaştırma, hasar (provizyon) verisinden gerçekleşen ortalama
  maliyet, hekim bazlı fiyat farkları.
- **Örnek veri:** Kurumlar, işlem kodları, fiyatlar ve adetler kurgusaldır.

## Testler

- Ortanca (tek ve çift sayıda kurum), işlem ve kurum endeksi, adet ağırlığı, yayılım, ortanca üstü tutar.
- Eşik sınırları (tam eşikteki endeks pahalı sayılmaz), aşırı fiyat, en az kurum koşulu, emsali olmayan kurum.
- Bölge ve grup ayrımı, mükerrer satır, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
