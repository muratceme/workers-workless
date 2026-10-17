# Teslimat Performansı (OTIF) Raporu · Kod Bloğu

> Lojistik ve Depo › Lojistik Uzmanı · Gerisi Bende

Siparişlerin **zamanında ve eksiksiz** (OTIF: On Time In Full) teslim oranını hesaplar; müşteri, taşıyıcı ve ay
bazında döker. Geç mi, eksik mi, ikisi birden mi olduğunu sipariş sipariş gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Zamanında (satır) | Teslim tarihi ≤ söz verilen tarih + `--tolerans` gün |
| Eksiksiz (satır) | Teslim miktarı ≥ sipariş miktarı × (1 − `--eksik-tolerans` ÷ 100) |
| Sipariş zamanında / eksiksiz | Siparişin **bütün** satırları zamanında / eksiksiz |
| OTIF | Sipariş hem zamanında hem eksiksiz |
| Zamanında %, Eksiksiz %, OTIF % | İlgili sipariş sayısı ÷ değerlendirilen sipariş sayısı × 100 |
| Miktar karşılama % | Σ min(teslim, sipariş) ÷ Σ sipariş miktarı × 100 |
| Açık | Teslim edilmemiş ve söz tarihi henüz gelmemiş satırı olan sipariş; oranlara girmez |
| Teslim edilmemiş, söz tarihi geçmiş | Hem geç hem eksik sayılır; gecikme durum tarihine kadar işler |

**Örnek:** Değerlendirilen 11 siparişin 6'sı hem zamanında hem eksiksiz → OTIF **%54,55**. Zamanında 7 sipariş
(%63,64), eksiksiz 8 sipariş (%72,73): OTIF her ikisinden de düşüktür, çünkü iki koşulu birlikte arar.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --siparisler siparisler.xlsx
python main.py --siparisler siparisler.xlsx --tolerans 1 --eksik-tolerans 2 --hedef 90
```

Her satır bir sipariş kalemidir; aynı sipariş numaralı satırlar tek sipariş olarak değerlendirilir. Zorunlu
sütunlar: Söz Verilen Tarih, Sipariş Miktarı. Teslim tarihi dolu, teslim miktarı boşsa tam teslim varsayılır.

## Çıktı

`cikti/otif_raporu.xlsx`:

- `Özet`: oranlar, sonuç dağılımı (OTIF, Geç, Eksik, Geç ve eksik, Açık).
- `Müşteriler`, `Taşıyıcılar`, `Aylar`: aynı oranlar kırılımlı; OTIF hedefin altındaysa kırmızı.
- `Siparişler`: her siparişin sonucu, gecikme günü, sipariş ve teslim miktarı.
- `Uyarılar`: söz tarihi geçtiği hâlde teslim edilmemiş siparişler, fazla teslim, okunamayan satırlar.

## Dikkat

- **OTIF tanımı şirketten şirkete değişir.** Burada sipariş bazında ve "bütün satırlar" kuralıyla hesaplanır; satır
  ya da miktar bazında ölçen bir müşteriyle sonuçlar tutmaz. Müşterinizle sözleşmede hangi tanım varsa ona göre
  toleransları ayarlayın; miktar bazlı görünüm için "Miktar karşılama %" sütununa bakın.
- **"Söz verilen tarih" hangi tarih?** Müşterinin istediği tarih mi, sizin teyit ettiğiniz tarih mi olduğu sonucu
  değiştirir. Bütün satırlarda aynı tanımı kullanın. Teslim tarihi de aynı şekilde: araca yükleme değil, müşteriye
  teslim tarihi olmalıdır.
- **Erken teslim zamanında sayılır.** Randevulu teslimatta erken gelen araç da sorun olabilir; teslim penceresi
  (en erken – en geç) kontrolü bu pakette yoktur.
- **Taşıyıcı sayfası sorumluluğu göstermez.** Geç teslimin nedeni üretim, stok ya da depo gecikmesi olabilir;
  taşıyıcıyı değerlendirmek için sevk tarihini ayrıca izlemek gerekir.
- **Az siparişte oran yanıltır.** Tek siparişi olan müşteride OTIF %0 ya da %100 çıkar; sipariş sayısıyla okuyun.
- **OTIF hedefi (%95) yalnız renklendirme içindir**; bir sektör standardı ya da öneri değildir.
- **Örnek veri:** Müşteriler, taşıyıcılar ve siparişler kurgusaldır.

## Testler

- Zamanında ve eksiksiz sınırları (aynı gün, bir gün geç, gün toleransı, yüzde toleransı), erken ve fazla teslim.
- Çok satırlı sipariş (tek satır geç ya da eksik), açık sipariş, söz tarihi geçmiş teslimsiz sipariş ve gecikmesi.
- Oranlar, miktar karşılama, müşteri / taşıyıcı / ay kırılımı, okuma ve komut satırı.

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
