# Nakliye Maliyet Analizi · Kod Bloğu

> Lojistik ve Depo › Lojistik Uzmanı · Gerisi Bende

Nakliye faturalarınızdan **kg, km ve ton-km başına taşıma maliyetini** hesaplar; bölge, taşıyıcı ve ay bazında
karşılaştırır, aynı bölgede pahalıya gelen seferleri ve düşük doluluğu işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Kg başı maliyet | Fatura tutarı ÷ taşınan ağırlık (kg) |
| Km başı maliyet | Fatura tutarı ÷ mesafe (km) |
| Ton-km başı maliyet | Fatura tutarı ÷ (ağırlık ÷ 1.000 × km) |
| Doluluk % | Taşınan ağırlık ÷ araç kapasitesi × 100 |
| Grup ortalaması | Toplam tutar ÷ toplam ağırlık (ağırlıklı ortalama) |
| Bölge ortalamasına fark % | (Taşıyıcının o bölgedeki kg maliyeti ÷ bölgenin kg maliyeti − 1) × 100 |
| Pahalı sefer | Kg maliyeti bölge medyanının `--sapma` (%30) üstünde; bölgede en az `--asgari` (3) sefer varsa |

**Örnek:** 8.000 kg için 20.000 → kg başı **2,50**. Aynı bölgenin medyanı 1,50; %30 üstü 1,95 → sefer işaretlenir.
Kapasite 20.000 kg → doluluk **%40**: pahalılığın nedeni yarı boş araç.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri
python main.py --seferler seferler.xlsx
python main.py --seferler seferler.xlsx --sapma 20 --doluluk 80 --asgari 5
```

Zorunlu sütunlar: Ağırlık ve Fatura Tutarı. Km yoksa km ve ton-km maliyeti, kapasite yoksa doluluk boş kalır.

## Çıktı

`cikti/nakliye_maliyet_analizi.xlsx`:

- `Bölgeler`, `Taşıyıcılar`, `Aylar`: sefer sayısı, ağırlık, km, tutar, birim maliyetler, sefer başı tutar, doluluk.
- `Bölge-Taşıyıcı`: aynı bölgede taşıyıcıların kg maliyeti ve bölge ortalamasına farkı.
- `Seferler`: her seferin birim maliyetleri, doluluğu ve bölge medyanı.
- `Uyarılar`: mükerrer sefer numarası, pahalı sefer, kapasite aşımı, düşük doluluk.

## Dikkat

- **Taşıyıcıları yalnız aynı bölgede karşılaştırın.** Uzun mesafede kg maliyeti doğal olarak yüksektir; taşıyıcı
  toplamları, hangi bölgelere çalıştıklarına göre değişir. Adil karşılaştırma için `Bölge-Taşıyıcı` sayfasına bakın.
- **Kg maliyeti doluluğa çok duyarlıdır.** Araç bazlı (komple) fiyatlanan seferde yarım yük kg maliyetini ikiye
  katlar; işaretlenen seferi doluluk sütunuyla birlikte okuyun. Hacimli ama hafif yüklerde ağırlık yerine hacim ya
  da palet sayısı daha doğru ölçüdür; bu paket yalnız ağırlığa bakar.
- **Tutarlar KDV hariç ve tek para biriminde olmalı.** Yakıt farkı, bekleme, indirme-bindirme gibi ek kalemleri
  tutara dahil edip etmediğiniz bütün satırlarda aynı olmalıdır.
- **Mükerrer numara her zaman çift fatura değildir.** Aynı numaralı ikinci satır hesaba katılmaz ve uyarılır; kısmi
  fatura ya da düzeltme ise numarayı ayırın.
- **Kapasite aşımı uyarısı veri kontrolüdür.** Aracın yasal azami yüklü ağırlığı ve dingil ağırlıkları bu paketin
  kapsamında değildir.
- **Kapsam dışı:** rota en iyileme, dönüş yükü, teslimat süresi ve hasar performansı, sözleşme fiyatıyla fatura
  mutabakatı.
- **Örnek veri:** Taşıyıcılar, seferler ve tutarlar kurgusaldır.

## Testler

- Birim maliyetler, ağırlıklı grup ortalaması, km ya da kapasitesi boş seferlerin toplamlara etkisi.
- Medyan (tek ve çift sayıda sefer), sapma eşiği sınırı, asgari sefer koşulu, doluluk sınırları ve kapasite aşımı.
- Mükerrer sefer numarası, bölge-taşıyıcı farkı, ay kırılımı, okuma ve komut satırı.

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
