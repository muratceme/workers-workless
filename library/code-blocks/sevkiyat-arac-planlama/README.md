# Sevkiyat Araç Planlama · Kod Bloğu

> Lojistik ve Depo › Lojistik Uzmanı · Gerisi Bende

Siparişleri teslim tarihi ve bölgeye göre gruplar; **palet, ağırlık ve hacim** kapasitelerine göre araçlara yerleştirir,
araç tipini seçer ve doluluk planı çıkarır. Araç kapasiteleri ve sefer maliyetleri sizin dosyanızdan okunur.
İnternete bağlanmaz.

## Nasıl hesaplar?

1. **Grup:** Aynı teslim tarihi ve aynı bölgedeki siparişler birlikte planlanır. Farklı gün veya bölge birleştirilmez.
2. **Bölme:** Hiçbir araca tek başına sığmayan sipariş, en çok yük alan aracın kapasitesi kadar parçalara bölünür.
   Palet tam sayıya yuvarlanır; ağırlık ve hacim aynı oranda paylaştırılır.
3. **Yerleştirme:** Siparişler büyükten küçüğe sıralanır. Her adımda her araç tipi için bir araç doldurulur; taşınan yük
   başına maliyeti en düşük olan tip seçilir.
4. **Küçültme:** Her sefer, yükünü taşıyabilen en ucuz araç tipine indirilir.
5. **Karşılaştırma:** Aynı grup yalnız tek araç tipiyle de denenir. Toplam maliyeti en düşük plan seçilir; eşitlikte
   sefer sayısı az olan.

| Terim | Hesap |
|---|---|
| Doluluk % | Palet, ağırlık ve hacim doluluklarının **en yükseği** |
| Sınırlayan | Doluluğu belirleyen boyut (araç neden doldu?) |
| Palet başına maliyet | Sefer maliyeti ÷ taşınan palet |

- Araç dosyasında "Mevcut Adet" girilirse o tipten günde en çok o kadar sefer planlanır. Boşsa sınırsız sayılır.
- Kapasitesi boş bırakılan boyut o araç için kontrol edilmez. Siparişte boş bırakılan boyut da kontrole katılmaz
  (uyarı yazılır).
- Doluluğu `--min-doluluk` (varsayılan %60) altında kalan seferler için uyarı verilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --siparisler siparisler.xlsx --araclar araclar.xlsx --min-doluluk 60
```

| Dosya | Sütunlar |
|---|---|
| Siparişler | Sipariş No, Müşteri, Teslim Bölgesi, Teslim Tarihi, Palet, Ağırlık (kg), Hacim (m3) |
| Araçlar | Araç Tipi, Palet Kapasitesi, Ağırlık Kapasitesi (kg), Hacim Kapasitesi (m3), Sefer Maliyeti, Mevcut Adet |

## Çıktı

`cikti/arac_plani.xlsx`:

- `Araç Planı`: sefer, tarih, bölge, araç tipi, yük ve kapasiteler, doluluk, sınırlayan boyut, maliyet.
- `Yükleme Listesi`: hangi sipariş hangi seferde (bölünen siparişlerde parça numarası).
- `Özet`: araç tipine göre sefer sayısı, ortalama doluluk ve toplam maliyet.
- `Uyarılar`: bölünen siparişler, düşük doluluk, atanamayan siparişler, veri hataları.

## Dikkat

- **Sezgisel yöntemdir.** Makul bir plan üretir, en düşük maliyeti garanti etmez. Sonucu bir planlamacı gözden
  geçirmelidir.
- **Kapsam dışı:** rota ve durak sırası, yükleme sırası, paletlerin üst üste konup konamayacağı, dingil ağırlığı,
  ADR / soğuk zincir gibi araç kısıtları, sürücü çalışma süreleri.
- **Sefer maliyeti** bölgeden bağımsız tek tutar olarak okunur. Bölgeye göre değişiyorsa her bölge için ayrı çalıştırın.
- **Karayolu ağırlık sınırları:** Ağırlık kapasitesine aracın ruhsatındaki istiap haddini yazın; azami yüklü ağırlık
  ve dingil sınırları için Karayolları Trafik Yönetmeliği geçerlidir.
- **Örnek veri:** Araç kapasiteleri ve maliyetler örnektir; kendi filonuzun değerlerini girin. Müşteri adları kurgusaldır.

## Testler

- Tek sipariş ve doluluk hesabı, küçültme, tek tipin daha ucuz olduğu durum.
- Gün ve bölge ayrımı, sipariş bölme (33 / 40 oranı), ağırlık sınırı, günlük araç adedi.
- Düşük doluluk uyarısı; okuma hataları; örnek veri uçtan uca (tüm paletler yüklü, kapasite aşımı yok).

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
