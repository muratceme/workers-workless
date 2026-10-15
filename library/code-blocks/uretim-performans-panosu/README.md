# Üretim Performans Panosu · Kod Bloğu

> Üretim › Üretim Müdürü · Workers / Workless

Günlük üretim kayıtlarından **OEE, plan uyumu, fire ve iş kazası** göstergelerini hat ve hafta bazında tek panoda
toplar; hedefin dışında kalan hücreleri renklendirir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Çalışma süresi | Planlı süre − duruş |
| Kullanılabilirlik % | Σ çalışma süresi ÷ Σ planlı süre × 100 |
| Performans % | Σ (ideal çevrim süresi × üretilen adet) ÷ Σ çalışma süresi × 100 |
| Kalite % | Σ sağlam adet ÷ Σ üretilen adet × 100 |
| OEE % | Kullanılabilirlik × performans × kalite |
| Plan uyumu % | Σ sağlam adet ÷ Σ planlanan adet × 100 |
| Fire % | Σ (üretilen − sağlam) ÷ Σ üretilen × 100 |
| Hafta | Tarihin ISO hafta numarası (hafta pazartesi başlar) |

Hat ve hafta oranları günlük oranların ortalaması değil, **toplamların oranıdır**. `Sağlam Adet` yoksa
üretilen − `Fire Adet` kullanılır. İdeal çevrim süresi saniye, süreler dakika cinsindendir.

Uyarılar hat × hafta düzeyinde üretilir: iş kazası kaydı (yüksek), OEE `--oee` (60) altında, plan uyumu `--plan`
(90) altında, fire `--fire` (3) üstünde, performans %100'ün üstünde.

**Örnek:** Planlı süre 450 dk, duruş 45 dk → çalışma 405 dk, kullanılabilirlik **%90**. İdeal çevrim 30 sn,
üretilen 720 adet → ideal süre 360 dk, performans **%88,89**. Sağlam 702 adet → kalite **%97,5**.
OEE = 0,90 × 0,8889 × 0,975 = **%78**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek veri
python main.py --kayitlar gunluk_uretim.xlsx
python main.py --kayitlar gunluk_uretim.xlsx --oee 70 --plan 95 --fire 2
```

Zorunlu sütunlar: Tarih, Hat, Planlı Süre (dk), İdeal Çevrim Süresi (sn), Üretilen Adet. Aynı hat ve gün için birden
çok satır (vardiya, ürün) olabilir; hepsi toplanır.

## Çıktı

`cikti/uretim_performans_panosu.xlsx`:

- `Pano (Hat × Hafta)`: bütün göstergeler; hedef dışı hücreler kırmızı.
- `Hatlar`, `Haftalar`: aynı göstergeler tek boyutta; hat sayfasının son satırı toplam.
- `Günlük Kayıtlar`: her satırın OEE değeriyle birlikte dökümü.
- `Uyarılar`.

## Dikkat

- **OEE, ideal çevrim süresi kadar doğrudur.** Çevrim süresi ürüne göre değişiyorsa her ürünü ayrı satırda kendi
  çevrim süresiyle yazın. Süre gerçek kapasiteden uzun tanımlanmışsa performans %100'ü aşar; kod bunu uyarır ama
  düzeltmez.
- **Planlı süre, planlı duruşları içermemelidir.** Mola, planlı bakım ve üretim planlanmayan süreleri planlı
  süreden düşerek yazın; `Duruş` sütununa yalnız plansız duruşları ve ayar sürelerini koyun. Farklı tanım,
  kullanılabilirliği olduğundan yüksek ya da düşük gösterir.
- **Hedefler sizin politikanızdır.** %60 OEE varsayılanı bir başlangıç eşiğidir; sık anılan "dünya sınıfı %85"
  değeri her süreç ve sektör için geçerli bir ölçüt değildir. Hatlarınızı kendi geçmişleriyle karşılaştırın.
- **Farklı ürün üreten hatların toplamı dikkatle okunmalı.** Toplam satırındaki kalite ve fire oranı adetleri,
  birim değeri çok farklı ürünleri eşit sayar.
- **Plan uyumu fazla üretimi ödüllendirir.** Plandan fazla üretim oranı %100'ün üstüne çıkarır; iş emri bazında
  sapma için ayrı bir plan uyum raporu kullanın.
- **İş kazası sayısı yalnız sizin yazdığınızdır.** Kaza sıklık ve ağırlık oranları, ramak kala olaylar ve yasal
  bildirimler kapsam dışıdır.
- **Örnek veri:** Hatlar, süreler ve adetler kurgusaldır.

## Testler

- Kullanılabilirlik, performans, kalite ve OEE; toplamların oranı olarak hat, hafta ve genel toplam.
- Fire adedinden sağlam adet türetme, plan uyumu, ISO hafta ataması (yıl sonu dahil).
- Hedef sınırları, %100'ü aşan performans, iş kazası uyarısı, tutarsız satırlar, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
