# İkinci El Araç Fiyat Değerleme · Kod Bloğu

> Satış › İkinci El Satış Danışmanı · Gerisi Bende

Topladığınız benzer ilan verisinden değerlenecek araç için **piyasa fiyat aralığını ve alım fiyatı önerisini**
üretir. İlanları siz sağlarsınız; kod internete bağlanmaz ve ilan sitelerinden veri çekmez.

## Nasıl hesaplar?

| Adım | Kural | Varsayılan |
|---|---|---|
| Emsal ilan | Aynı marka ve model; yıl farkı ≤ `--yil-fark`; km ±`--km-fark` %; yakıt ve vites yazılıysa aynı | 1 · %30 |
| Az emsal | Sayı `--asgari` altındaysa km koşulu kaldırılır; yine azsa "Yetersiz veri" | 3 |
| Aykırı fiyat | En az 4 ilan varsa Q1 − 1,5 × ÇAA altı ve Q3 + 1,5 × ÇAA üstü çıkarılır | |
| Piyasa değeri | Kalan fiyatların medyanı; aralık alt çeyrek – üst çeyrek | |
| Alım önerisi | Medyan × (1 − `--marj` %), `--yuvarla` katına aşağı yuvarlanır | %10 · 1.000 |

ÇAA (çeyrekler arası aralık) = Q3 − Q1. Çeyrekler doğrusal ara değerle hesaplanır.

**Örnek:** 2020 model, 85.000 km. Emsal fiyatlar 820.000, 880.000, 900.000, 950.000 ve 2.000.000. Sonuncusu aykırı
olduğu için çıkarılır. Medyan 890.000; aralık 865.000 – 912.500; %10 marjla alım önerisi 801.000.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --ilanlar ilanlar.xlsx --araclar araclar.xlsx
python main.py --ilanlar ilanlar.xlsx --araclar araclar.xlsx --yil-fark 2 --km-fark 20 --marj 12 --asgari 5
```

- İlanlar: Marka, Model, Yıl, Fiyat (zorunlu); Km, Yakıt, Vites, Paket.
- Araçlar: Marka, Model, Yıl (zorunlu); Araç (plaka veya ad), Km, Yakıt, Vites, Paket, Hasar Tutarı.

## Çıktı

`cikti/ikinci_el_degerleme.xlsx`:

- `Değerleme`: araç bazında emsal sayısı, en düşük, alt çeyrek, medyan, üst çeyrek, en yüksek, alım önerisi.
- `Emsal İlanlar`: her araç için hesaba giren ilanlar.
- `Uyarılar`: yetersiz veri, kaldırılan km koşulu, hasar kaydı, az sayıda emsal, okunamayan satırlar.

## Dikkat

- **İlan fiyatı satış fiyatı değildir.** İlanlar istenen fiyatı gösterir; gerçek satışlar genellikle daha düşüktür.
  Marj bu farkı, hazırlık giderini ve kârınızı karşılamak içindir; kendi deneyiminize göre ayarlayın.
- **Araç durumu hesaba girmez.** Hasar kaydı, boya, değişen parça, bakım geçmişi, lastik ve donanım farkını kod
  bilmez; öneriyi ekspertiz sonucuna göre düzeltin. Hasar tutarı yazılan araçlar yalnız uyarılır.
- **Paket eşleştirilmez.** Donanım paketi bilgi olarak taşınır; aynı modelin giriş ve üst paketleri arasındaki fiyat
  farkı aralığı genişletir.
- **Az ilanla sonuç güvenilmezdir.** Beşten az emsalde aralık dar veya yanıltıcı olabilir.
- **Veri toplama sizin sorumluluğunuzdadır.** İlan sitelerinin kullanım koşullarına ve kişisel verilerin korunması
  mevzuatına uygun davranın.
- **Örnek veri:** Markalar, modeller ve fiyatlar kurgusaldır.

## Testler

- Çeyrek hesabı, emsal seçimi, aykırı fiyatın çıkarılması, medyan, aralık ve alım önerisi.
- Km koşulunun kaldırılması, yetersiz veri, km'si olmayan araç, yakıt ve vites eşleşmesi, marj ve yuvarlama ayarları.
- Okunamayan ilan ve araç satırları, komut satırı hataları; örnek veri uçtan uca.

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
