# Rakip Analizi Raporu · Kod Bloğu

> Pazarlama › Pazarlama Uzmanı · Workers / Workless

Rakip fiyat gözlemlerini, kampanyalarını ve ürün özelliklerini ürünlerimizle karşılaştıran bir rapor hazırlar.
Fiyatlar **birim fiyat endeksiyle** karşılaştırıldığı için farklı ambalaj boyları (ör. 2,5 L ve 3 L) doğru
kıyaslanır. İnternete bağlanmaz; fiyat ve kampanya verileri kullanıcı tarafından toplanır (mağaza ziyareti,
e-ticaret siteleri, broşürler).

## Nasıl hesaplar?

- **Birim fiyat:** fiyat / ambalaj miktarı.
  - Miktarlar temel birime çevrilir: kg → g, L → ml, cl → ml.
  - İndirimli fiyat girilmişse ve liste fiyatından düşükse o kullanılır.
- **Endeks:** rakibin birim fiyatı / bizim birim fiyatımız × 100.
  - 100'ün üstü rakibin daha pahalı olduğunu gösterir (yeşil ≥ 105). 100'ün altı rakibin daha ucuz olduğunu gösterir
    (kırmızı ≤ 95).
  - Her rakip-ürün için en son gözlem kullanılır.
  - Birimler uyuşmazsa (ör. g ve ml) veya miktar eksikse paket fiyatları karşılaştırılır ve satıra not düşülür.
- **Konum:**
  - Tüm rakipler pahalıysa "En ucuz", tüm rakipler ucuzsa "En pahalı", karışıksa "Pazar ortası".
  - Rakiplerin medyan endeksi kullanılır, böylece tek bir özel marka sonucu bozmaz.
- **Fiyat değişimi:** Aynı rakip ürününün ilk ve son gözlemi karşılaştırılır.
- **Kampanyalar:** Rakip bazında sayı ve ortalama indirim gösterilir; rapor tarihinde süren kampanyalar işaretlenir.
- **Özellik matrisi:**
  - "Eksik": bizde yok, rakiplerin en az yarısında var.
  - "Avantaj": yalnız bizde var.

| Uyarı | Önem |
|---|---|
| Fiyat dezavantajı: bizim birim fiyat rakiplerin medyanından `--esik` (%10) fazla yüksek | Orta |
| Süren derin indirim: bizim kategorilerimizde `--derin` (%30) ve üzeri | Orta |
| Fiyat bırakma olasılığı: tüm rakipler bizden %10'dan fazla pahalı | Bilgi |
| Özellik eksiği; eşleşmeyen fiyat satırı (ürün kodu boş / yok) | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                      # örnek: 12 ürün, 4 rakip, rapor tarihi 09.10.2026
python main.py --urunler urunler.xlsx --fiyatlar rakip_fiyatlari.xlsx --kampanyalar kampanyalar.xlsx --ozellikler ozellikler.xlsx
```

| Dosya | Sütunlar |
|---|---|
| Ürünler | Ürün Kodu, Ürün, Kategori, Fiyat, Ambalaj Miktarı, Birim (g, kg, ml, L, adet...) |
| Rakip fiyatları | Tarih, Rakip, Kanal, Bizim Ürün Kodu (karşılaştırılan ürünümüz), Rakip Ürün, Fiyat, İndirimli Fiyat, Ambalaj Miktarı, Birim, Stokta (Evet / Hayır) |
| Kampanyalar (isteğe bağlı) | Başlangıç, Bitiş, Rakip, Kanal, Kampanya, İndirim (%), Kategori (boşsa tüm ürünler) |
| Özellikler (isteğe bağlı) | Marka (bizim için `Biz`), Kategori, Özellik, Değer (Var / Yok veya serbest değer) |

## Çıktı

`rakip_analizi.xlsx`:
- `Özet`.
- `Fiyat Karşılaştırma`: ürün × rakip endeksi, medyan endeks, konum.
- `Rakip Özeti`: medyan endeks, bizden ucuz / pahalı ürün sayısı, stokta olmama ve indirimli gözlem oranı, kampanya
  özeti.
- `Kategori Özeti`, `Fiyat Değişimi`, `Kampanyalar`.
- `Özellik Matrisi`: eksik ve avantaj olan özellikler renkli.
- `Veri`, `Uyarılar`.

## Dikkat

- **Eşleştirme kullanıcıdadır.** "Bizim Ürün Kodu" sütunu, rakip ürününün hangi ürünümüzle karşılaştırılacağını
  belirler. Kalite ve içerik farkını göz önüne alarak eşleştirin.
- **Fiyat kararları** yalnız endekse göre verilmemelidir. Maliyet, marj, marka konumu ve rekabet hukuku dikkate
  alınmalıdır; rakiplerle fiyat bilgisi paylaşımı ve fiyat koordinasyonu 4054 sayılı Kanun'a aykırıdır.
- **Veri toplama:** Rakip sitelerinden otomatik veri çekmeden önce sitenin kullanım koşullarını kontrol edin.
- **Örnek veri:** Örnek ürünler, rakipler ve fiyatlar kurgusaldır.

## Testler

Şunlar test edilir:
- Farklı ambalaj boyunda birim fiyat endeksi (elle hesap), son gözlem ve indirimli fiyat.
- Konum, fiyat dezavantajı ve bırakma, derin indirim eşiği, eşleşmeyen satır.
- Özellik eksiği ve avantajı, fiyat değişimi, süren kampanyalar.
- Miktar okuma, birim uyuşmazlığı, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
