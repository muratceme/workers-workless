# Fiyat Farkı Hesaplama · Kod Bloğu

> Teknik Ofis › Teknik Ofis Şefi · Gerisi Bende

Sözleşme katsayılarına (a, b1, b2, c...) ve endekslere göre **dönem fiyat farkı katsayısını (Pn) ve tutarını**
hesaplar. İnternete bağlanmaz; endeks değerlerini siz verirsiniz.

## Nasıl hesaplar?

```
Pn = a × (In ÷ Io) + b1 × (Çn ÷ Ço) + b2 × (Dn ÷ Do) + ... + c × (Mn ÷ Mo)
F  = An × B × (Pn − 1)
```

| Simge | Anlamı |
|---|---|
| Katsayılar | Sözleşmede yazan ağırlıklar; toplamları 1,00 olmalıdır (değilse hesap yapılmaz) |
| …o | Endeksin temel aydaki değeri (`--temel-donem`) |
| …n | Endeksin uygulama ayındaki değeri |
| An | Uygulama ayında sözleşme fiyatlarıyla yapılan iş tutarı |
| B | `--b`, varsayılan 0,90 |

- Katsayı sayısı ve adları serbesttir: katsayı dosyasındaki her satır kendi endeksine bağlanır.
- Pn altı ondalık haneye, fiyat farkı kuruşa yuvarlanır. Pn < 1 ise fiyat farkı eksi çıkar (kesinti).
- Uygulama ayının endeksi dosyada yoksa o hakediş hesaplanmaz ve uyarılır.

**Örnek:** katsayılar 0,30 / 0,25 / 0,25 / 0,20; endeks oranları 1,10 / 1,05 / 1,10 / 1,00 → Pn = 0,33 + 0,2625 + 0,275 +
0,20 = 1,0675. An 2.000.000 → F = 2.000.000 × 0,90 × 0,0675 = 121.500.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (temel ay 2026-03)
python main.py --katsayilar katsayilar.xlsx --endeksler endeksler.xlsx --hakedisler hakedisler.xlsx --temel-donem 2026-03
```

| Dosya | Sütunlar |
|---|---|
| Katsayılar | Katsayı (a, b1 …), Değer, Endeks (endeks dosyasındaki adla aynı) |
| Endeksler | Dönem, Endeks, Değer. Temel ay ve her uygulama ayı için birer satır |
| Hakedişler | Dönem (uygulama ayı), Tutar (An) (zorunlu); Hakediş No |

Dönem: `2026-06`, `06.2026` veya `Haziran 2026`.

## Çıktı

`cikti/fiyat_farki.xlsx`:

- `Fiyat Farkı`: hakediş başına An, Pn, fiyat farkı ve kümülatif toplam.
- `Pn Detayı`: her ay için katsayı, temel ve uygulama ayı endeksi, endeks oranı ve Pn'e katkısı.
- `Uyarılar`: endeksi olmayan aylar, okunamayan satırlar.

## Dikkat

- **Kaynak teyidi yapılmamıştır.** Formül, kamu yapım işlerinde uygulanan fiyat farkı esaslarının genel yapısına göre
  kurulmuştur. Sözleşmenizin tabi olduğu esasları, B katsayısını, yuvarlama kuralını ve temel ayın nasıl belirlendiğini
  (ihale veya teklif tarihine göre) kendi mevzuat ve sözleşme metninizden doğrulayın.
- **Endeksler sizin verinizdir.** Kod TÜİK'ten veri indirmez; hangi endeks tablosunun ve hangi satırın kullanılacağı
  sözleşmede yazar. Örnek dosyadaki endeks adları ve değerleri kurgusaldır.
- **Kapsam dışı durumlar:** avans mahsubu, süre uzatımı ve gecikme halinde hangi ayın endeksinin kullanılacağı, iş
  programına göre ödenek dilimleri, güncel fiyatla yapılan işler, sabit katsayılı veya ek fiyat farkı düzenlemeleri
  hesaplanmaz. Bu durumlarda An tutarını ve endeks ayını kendiniz düzeltip girin.
- **İdarenin hesabı esastır.** Sonuç ön hesap ve kontrol amaçlıdır.
- **Örnek veri:** Katsayılar, endeksler ve hakediş tutarları kurgusaldır.

## Testler

- Pn ve fiyat farkı, katsayı katkıları, kümülatif toplam, düşen endekste eksi fiyat farkı, yuvarlama.
- Endeksi olmayan ay, toplamı 1 olmayan katsayılar, temel ayı olmayan endeks; dönem ve sayı okuma.
- Okuma hataları ve örnek veri uçtan uca.

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
