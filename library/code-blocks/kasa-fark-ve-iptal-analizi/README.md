# Kasa Fark ve İptal Analizi · Kod Bloğu

> Mağaza Operasyonları › Kasa Sorumlusu · Workers / Workless

Kasiyer bazında **iade, iptal ve kasa farklarını** dönem boyunca toplar; oranı diğer kasiyerlerden belirgin yüksek
olan, art arda eksik veren ya da sık fark çıkan kasiyerleri **inceleme için** işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| İade %, iptal % | İade ya da iptal tutarı ÷ satış tutarı × 100 |
| 100 fişte iade / iptal | İade ya da iptal adedi ÷ fiş adedi × 100 |
| Farklı gün | Kasa farkının mutlak değeri `--tolerans` (5) üstünde olan gün |
| Eksik gün, fazla gün | Farkı toleransın altında eksi ya da üstünde artı olan günler |
| En uzun eksik serisi | Kasiyerin çalıştığı günler sırasıyla, art arda eksik verdiği gün sayısı |
| Net kasa farkı | Günlük farkların toplamı; eksik ve fazlalar birbirini götürür |
| Mutlak fark toplamı | Günlük farkların mutlak değerlerinin toplamı |

Dönemde fiş adedi `--min` (100) ve üstünde olan kasiyerler için işaretler:

| İşaret | Kural |
|---|---|
| İade oranı yüksek | Kasiyerin iade %'si, **diğer kasiyerlerin** toplam iade %'sinin `--kat` (2) katından büyük |
| İptal oranı yüksek | Aynı karşılaştırma iptal için |
| Sürekli kasa eksiği | En uzun eksik serisi `--seri` (3) gün ve üstü |
| Sık kasa farkı | Farklı gün ÷ çalışılan gün × 100, `--fark-gun` (50) üstünde |

Kasiyer kendi oranıyla karşılaştırılmaz: karşılaştırma değeri, o kasiyer dışındaki herkesin toplamından hesaplanır.

**Örnek:** Bir kasiyerin iade oranı %7; diğer kasiyerlerin toplam iade oranı %1,43, iki katı %2,86 → "İade oranı
yüksek". Aynı kasiyer art arda 4 gün eksik vermiş → "Sürekli kasa eksiği".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                # örnek veri
python main.py --kayitlar kasiyer_gunluk.xlsx
python main.py --kayitlar kasiyer_gunluk.xlsx --tolerans 10 --kat 1,5 --seri 4 --fark-gun 40 --min 500
```

Zorunlu sütunlar: Tarih, Kasiyer, Satış Tutarı. Her satır bir kasiyerin bir günüdür.

## Çıktı

`cikti/kasa_fark_iptal_analizi.xlsx`:

- `Kasiyerler`: en çok işaret alandan başlayarak; değerlendirme ve bütün göstergeler; son satır mağaza toplamı.
- `Günlük Kayıtlar`: kasiyer ve tarih sırasıyla; farklı günler renkli.
- `Uyarılar`.

## Dikkat

- **İşaret, suistimal tespiti değildir.** Yüksek iade ya da iptal oranının meşru nedenleri olabilir: iade kasasında
  çalışmak, yoğun vardiya, yeni başlamış olmak, sorunlu bir ürün ya da kampanya. Rapor nereye bakılacağını gösterir;
  sonuç ancak fiş, onay kayıtları ve gerekiyorsa kamera incelemesiyle çıkarılır.
- **Kişisel veri içerir.** Rapor çalışan performansına ilişkindir; erişimi sınırlayın, saklama ve paylaşımda
  kişisel verilerin korunması kurallarına ve şirket içi usullere uyun. Çalışan hakkında işlem yapılmadan önce insan
  kaynakları ve hukuk birimleriyle birlikte değerlendirin.
- **Az kasiyerde karşılaştırma zayıftır.** İki-üç kasiyerli mağazada "diğerleri" tek kişidir; eşikleri buna göre
  yorumlayın.
- **Oranlar tutar üzerindendir.** Tek büyük iade oranı yükseltir; adet bazlı sütunlara (100 fişte iade / iptal) da
  bakın.
- **Kasa farkını siz girersiniz.** Kod sayım yapmaz; fark, gün sonu kasa mutabakatından gelmelidir.
- **Kapsam dışı:** indirim ve fiyat değiştirme işlemleri, çekmece açma sayıları, fiş bazında inceleme, saat bazlı
  örüntüler.
- **Örnek veri:** Kasiyerler ve tutarlar kurgusaldır.

## Testler

- Oranlar, farklı / eksik / fazla gün, en uzun eksik serisi (araya fark olmayan gün girince sıfırlanması).
- İşaret kuralları ve eşikleri, "diğer kasiyerler" karşılaştırması, en az fiş koşulu, tek kasiyerli dosya.
- Sıralama, uyarı düzeyi, okuma ve komut satırı.

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
