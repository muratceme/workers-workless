# Numune Ölçü Kontrolü · Kod Bloğu

> Modelhane › Modelist · Workers / Workless

Numune ölçümlerini **ölçü tablosu ve toleranslarla** karşılaştırır; tolerans dışı ölçü noktalarını, sapmanın yönünü
ve gereken düzeltme miktarını raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Fark         = ölçülen − ölçü tablosundaki beden değeri
Tolerans içi = −tolerans (−) ≤ fark ≤ +tolerans (+)        (sınır değer tolerans içindedir)
Düzeltme     = farkın tersi: fark +2,0 ise "2 küçült", −1,5 ise "1,5 büyüt"
```

| Numune sonucu | Koşul |
|---|---|
| Kritik ölçü tolerans dışı | `Kritik` işaretli en az bir ölçü tolerans dışında |
| Tolerans dışı | Yalnız kritik olmayan ölçüler tolerans dışında |
| Uygun | Tüm ölçüler tolerans içinde |

- **Ölçü tablosu geniş biçimdedir:** `Ölçü Noktası`, `Kritik`, tolerans ve ardından her beden için bir sütun.
- **Tolerans:** tek sütun `Tolerans (±)` veya ayrı `Tolerans (-)` ve `Tolerans (+)`. İşaret önemli değildir.
- **Model sütunu** tabloda boşsa satır tüm modellere uygulanır.
- "Ölçü Noktası Özeti" aynı noktanın bedenler arasında hep aynı yönde sapıp sapmadığını gösterir.
- Sayılarda virgül de nokta da ondalık ayırıcıdır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --olcumler numune_olcumleri.xlsx --tablo olcu_tablosu.xlsx
```

| Dosya | Sütunlar |
|---|---|
| Numune ölçümleri | Beden, Ölçü Noktası, Ölçülen (zorunlu); Model, Numune |
| Ölçü tablosu | Ölçü Noktası (zorunlu), Kritik, Tolerans (±) veya Tolerans (-) / Tolerans (+), Model; sonra beden sütunları |

Ölçü noktası adları iki dosyada aynı yazılmalıdır.

## Çıktı

`cikti/numune_olcu_kontrolu.xlsx`:

- `Numune Özeti`: numune bazında sonuç, tolerans dışı kritik ve diğer ölçüler; boş "Karar / Yorum" sütunu.
- `Ölçü Karşılaştırma`: her ölçü için tablo değeri, ölçülen, fark, tolerans, durum ve düzeltme.
- `Ölçü Noktası Özeti`: nokta bazında tolerans dışı sayısı, yön ve ortalama fark.
- `Uyarılar`: tabloda olmayan ölçü noktası veya beden, okunamayan değerler.

## Dikkat

- **Ölçüm yöntemi aynı olmalı.** Yarım / tam ölçü, ölçünün alındığı nokta ve ürünün serili durumu ölçü tablosundaki
  tanımla aynı değilse fark anlamsızdır. Birim (cm / inç) de aynı olmalıdır.
- **Yıkama ve çekme:** Ölçüm yıkama öncesi mi sonrası mı alındıysa tablo da o duruma ait olmalıdır.
- **Düzeltme miktarı kalıp farkı değildir.** Yarım ölçülerde, esneyen kumaşlarda veya çekme payı olan ürünlerde
  kalıba uygulanacak değer modelistin kararıdır.
- **Onay kararı değildir.** Numunenin onayı müşteri yorumlarına ve fit değerlendirmesine bağlıdır; "Kritik" işareti
  ölçü tablosunda sizin verdiğiniz bilgidir.
- **Örnek veri:** Model, ölçüler ve toleranslar kurgusaldır.

## Testler

- Tolerans sınırları (sınır değer içeride), farklı eksi / artı tolerans, düzeltme metni.
- Numune sonucu (kritik / kritik olmayan), ölçü noktası özeti.
- Modele özel ve genel tablo satırları, tabloda olmayan beden ve nokta; okuma hataları; örnek veri uçtan uca.

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
