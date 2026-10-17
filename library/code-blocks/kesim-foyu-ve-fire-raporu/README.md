# Kesim Föyü ve Fire Raporu · Kod Bloğu

> Üretim ve Fason Takip › Kesimhane Şefi · Gerisi Bende

Kesilen adetleri **beden-renk bazında sipariş adetleriyle** karşılaştırır; kumaş kullanımını, gerçekleşen birim
sarfiyatı ve **kesim firesini** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Kesim farkı                = kesilen adet − sipariş adedi
Kesim %                    = kesilen ÷ sipariş × 100
Kullanılan kumaş           = verilen − iade
Planlanan kullanım         = kesilen adet × planlanan birim sarfiyat
Kesim firesi               = kullanılan − planlanan kullanım
Fire %                     = kesim firesi ÷ kullanılan × 100
Gerçekleşen birim sarfiyat = kullanılan ÷ kesilen adet
```

| Durum | Koşul |
|---|---|
| Eksik kesim | Kesim % < 100 − `--eksik-tol` (varsayılan 0) |
| Fazla kesim | Kesim % > 100 + `--fazla-tol` (varsayılan 5) |
| Uygun | Arada |
| Fire yüksek | Fire % > `--fire-esik` (varsayılan 3) |

- Aynı sipariş-renk-beden için birden çok kesim satırı (ayrı pastallar) toplanır.
- Kumaş hesabı sipariş + renk bazındadır. Planlanan birim sarfiyat, pastal planından çıkan adet başı kumaştır.
- **Örnek:** 1.025 adet kesildi, 1.300 m verildi, 40 m iade → kullanılan 1.260 m. Planlanan sarfiyat 1,20 m/adet →
  planlanan kullanım 1.230 m → kesim firesi 30 m (%2,38); gerçekleşen sarfiyat 1,229 m/adet.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --siparis siparis.xlsx --kesim kesim.xlsx --kumas kumas.xlsx --fazla-tol 3 --fire-esik 2
```

| Dosya | Sütunlar |
|---|---|
| Sipariş | Sipariş No, Beden, Sipariş Adedi (zorunlu); Model, Renk |
| Kesim | Sipariş No, Beden, Kesilen Adet (zorunlu); Renk, Pastal No |
| Kumaş (isteğe bağlı) | Sipariş No, Verilen Kumaş (zorunlu); Renk, İade, Planlanan Birim Sarfiyat, Birim |

## Çıktı

`cikti/kesim_foyu.xlsx`:

- `Kesim Föyü`: sipariş-renk-beden bazında sipariş, kesilen, fark, kesim % ve durum.
- `Kumaş ve Fire`: renk bazında verilen, iade, kullanılan, planlanan kullanım, fire ve birim sarfiyat; boş "Neden" sütunu.
- `Uyarılar`: siparişte olmayan kesimler, kumaş kaydı olmayan renkler, okunamayan satırlar.

## Dikkat

- **Kesim firesi plana göre farktır.** Pastal planının kendi verimi (pastal içi fire) bu hesapta görünmez; yalnız
  plan ile gerçekleşen kullanım arasındaki fark (top sonu, pastal başı-sonu payı, hata kesimi, en farkı) ölçülür.
- **Birim tutarlı olmalı.** Verilen kumaş, iade ve birim sarfiyat aynı birimde (metre veya kg) girilmelidir; örme
  kumaşta kg / metre dönüşümünü siz yaparsınız.
- **Ortak kumaş:** Aynı toptan birden çok siparişe kesim yapıldıysa kumaşı siparişlere siz paylaştırmalısınız.
- **Toleranslar siparişe özeldir.** Müşterinin kabul ettiği eksik / fazla yükleme oranını `--eksik-tol` ve
  `--fazla-tol` ile girin; varsayılanlar genel bir kural değildir.
- **Kesilen adet sağlam adet değildir.** Kesim sonrası metolama ve dikim firesi ayrıca izlenmelidir.
- **Örnek veri:** Sipariş, kesim ve kumaş miktarları kurgusaldır.

## Testler

- Kesim durumu sınırları (%99, %105, %106), hiç kesilmemiş beden, eksik toleransı.
- Kumaş kullanımı, fire, birim sarfiyat, fire eşiği, plandan az kullanım, sarfiyatı olmayan renk.
- Pastal toplama, siparişte olmayan kesim, okuma hataları; örnek veri uçtan uca.

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
