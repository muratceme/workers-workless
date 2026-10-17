# Kampanya Etki Analizi · Kod Bloğu

> Satın Alma ve Kategori Yönetimi › Kategori Uzmanı · Gerisi Bende

Kampanya dönemindeki **ek satışı (uplift), kanibalizasyonu ve kampanya kârlılığını** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Beklenen adet | Baz dönem adedi × kampanya gün sayısı ÷ baz dönem gün sayısı |
| Ek satış, ek satış % | Kampanya dönemi adedi − beklenen adet; ek satış ÷ beklenen × 100 |
| Kâr farkı | Kampanya adedi × (uygulanan fiyat − birim maliyet) − beklenen adet × (normal fiyat − birim maliyet) |
| İndirim maliyeti | Beklenen adet × (normal fiyat − kampanya fiyatı) |
| Başabaş adet | Beklenen kâr ÷ (kampanya fiyatı − birim maliyet), yukarı yuvarlanır |
| Kanibalizasyon % | Aynı kategorideki kampanya dışı ürünlerin adet kaybı ÷ kampanya ürünlerinin ek satışı × 100 |
| Net etki | Kampanya ürünlerinin kâr farkı + aynı kategorideki diğer ürünlerin kâr farkı − kampanya gideri |

- Kampanya dışı ürünlerde adet artmışsa "Ek satış (halo)" olarak gösterilir ve net etkiye eklenir.
- İçinde kampanya ürünü olmayan kategoriler "Kapsam dışı" sayılır; listelenir ama toplamlara girmez.

**Örnek:** baz dönem 28 gün 400 adet, kampanya 7 gün 260 adet; normal fiyat 200, kampanya fiyatı 160, maliyet 120 →
beklenen 100, ek satış 160 (%160); kâr 260 × 40 = 10.400, beklenen kâr 100 × 80 = 8.000, kâr farkı +2.400; başabaş
8.000 ÷ 40 = 200 adet.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (28 gün baz, 7 gün kampanya, gider 1.000)
python main.py --satislar kampanya_satislari.xlsx --baz-gun 28 --kampanya-gun 7 --gider 15000
```

Sütunlar: Ürün, Baz Dönem Adet, Kampanya Dönemi Adet, Normal Fiyat, Birim Maliyet (zorunlu); Kategori, Kampanyada
(Evet / Hayır), Kampanya Fiyatı.

## Çıktı

`cikti/kampanya_etki.xlsx`:

- `Özet`: ek satış, ek ciro, indirim maliyeti, kâr farkları, gider ve net etki.
- `Ürün Bazında`: beklenen ve gerçekleşen adet, kâr farkı, kampanya ürünlerinde başabaş adet.
- `Kategori Bazında`: kanibalizasyon veya halo etkisi, kategori net kâr farkı.
- `Uyarılar`: okunamayan değerler, maliyetin altındaki kampanya fiyatı, baz dönemde satışı olmayan ürünler.

## Dikkat

- **Beklenen adet düz ortalamadır.** Mevsimsellik, bayram, maaş haftası, hava durumu ve trend hesaba katılmaz. Baz
  dönemi kampanyasız, stok sorunu olmayan ve kampanya dönemine benzeyen haftalardan seçin.
- **Kampanya sonrası düşüş yoktur.** Müşteri stok yaptıysa kampanyadan sonraki haftalarda satış düşer; bu kayıp
  hesapta yer almaz. Kampanya sonrası dönemi ayrıca izleyin.
- **Kanibalizasyon aynı kategoriyle sınırlıdır.** Başka kategorideki ikame ürünler ve diğer mağazalardan kayan satış
  görülmez. Diğer ürünlerdeki değişimin tamamı kampanyaya bağlanır; başka nedenleri olabilir.
- **Fiyat ve maliyet aynı esasta olmalı.** Hepsi KDV hariç olmalıdır. Tedarikçi kampanya desteği varsa birim
  maliyetten düşün veya gideri azaltın.
- **Örnek veri:** Kategoriler, ürünler ve tutarlar kurgusaldır.

## Testler

- Beklenen adet, ek satış, kâr farkı, indirim maliyeti ve başabaş adet; maliyetin altında fiyat, baz satışı olmayan ürün.
- Kanibalizasyon, halo, kapsam dışı kategori ve net etki; başabaş sonucu.
- Okuma hataları, zorunlu parametreler ve örnek veri uçtan uca.

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
