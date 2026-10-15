# Kumaş Fiyat Takibi · Kod Bloğu

> Kumaş ve Aksesuar Satın Alma › Kumaş Satın Alma Uzmanı · Workers / Workless

Tedarikçilerden gelen kumaş ve iplik fiyatlarını tek tabloda toplar; **kalite bazında fiyat geçmişini** ve
tedarikçilerin güncel tekliflerinin karşılaştırmasını üretir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Fiyat serisi | Kalite + tedarikçi + birim + para birimi; kayıtlar tarihe göre sıralanır |
| Son değişim % | (Son fiyat − önceki fiyat) ÷ önceki fiyat × 100 |
| Dönem değişimi % | (Son fiyat − ilk fiyat) ÷ ilk fiyat × 100 |
| Güncel teklif | Her serinin en son tarihli fiyatı |
| Kalite karşılaştırması | Aynı kalite, birim ve para birimindeki güncel teklifler: en düşük, ortanca, en yüksek |
| Fark % | (En yüksek − en düşük) ÷ en düşük × 100 |
| Yaş (gün) | Durum tarihi − son fiyat tarihi |

Uyarılar: son değişim `--artis` (10) yüzdesini aşan artış; aynı büyüklükte düşüş (bilgi); son fiyatı `--eski` (90)
günden eski teklif (bilgi). Birim boşsa `m`, para birimi boşsa `TRY` kabul edilir.

**Örnek:** Bir tedarikçinin süprem fiyatı 300'den 345'e çıktı → son değişim **%15**; dönem başındaki 285'e göre
**%21,05**. Aynı kalitede güncel teklifler 310, 320 ve 345 → en düşük 310, fark **%11,29**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --fiyatlar fiyat_listeleri.xlsx
python main.py --fiyatlar fiyat_listeleri.xlsx --artis 5 --eski 60 --tarih 31.10.2026
```

Zorunlu sütunlar: Tarih, Tedarikçi, Kalite, Fiyat. Yeni fiyat geldikçe dosyaya satır ekleyin; eski satırları silmeyin.

## Çıktı

`cikti/kumas_fiyat_takibi.xlsx`:

- `Kalite Karşılaştırması`: kalite başına güncel teklif sayısı, en düşük ve en yüksek fiyat ile tedarikçileri, ortanca, fark.
- `Güncel Fiyatlar`: tedarikçi bazında son fiyat, sıra, en düşüğe göre fark, son ve dönem değişimi, en düşük ve en yüksek.
- `Fiyat Geçmişi`: bütün kayıtlar, önceki kayda göre değişimle.
- `Uyarılar`.

## Dikkat

- **Aynı "kalite" aynı kumaş olmayabilir.** En, gramaj, iplik numarası, boyama türü, minimum sipariş, teslim şekli ve
  vade fiyatı değiştirir. Kod yalnız kalite adını eşleştirir; karşılaştırma ancak bu koşullar eşitse anlamlıdır.
  Farklı koşulları kalite adına ekleyin ("30/1 Süprem 180 gr, 60 gün vade" gibi).
- **Para birimleri çevrilmez.** TRY ve USD fiyatları ayrı gruplarda karşılaştırılır; kur etkisini hesaplamaz.
- **TRY fiyatlarda değişim enflasyonu da içerir.** Yüzde değişim reel artışı göstermez.
- **Eski teklif "güncel" görünür.** Bir tedarikçi uzun süredir fiyat vermediyse son bilinen fiyatı karşılaştırmaya
  girer; yaş sütununa ve "Fiyat eski" bildirimine bakın.
- **En düşük fiyat en iyi teklif demek değildir.** Kalite, termin ve tedarikçi güvenilirliği bu raporda yoktur.
- **Kalite ve tedarikçi adları tutarlı yazılmalı.** Büyük-küçük harf ve Türkçe karakter farkı yok sayılır; başka
  yazım farkları ayrı seri oluşturur.
- **Örnek veri:** Tedarikçiler, kaliteler ve fiyatlar kurgusaldır.

## Testler

- Son ve dönem değişimi, tek kayıtlı seri, aynı tarihte iki fiyat.
- Kalite karşılaştırması: en düşük, ortanca, en yüksek, fark; para birimi ve birime göre ayrım; tedarikçi sırası.
- Artış ve eski fiyat eşikleri, okuma ve komut satırı.

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
