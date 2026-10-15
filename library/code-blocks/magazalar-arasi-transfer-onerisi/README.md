# Mağazalar Arası Transfer Önerisi · Kod Bloğu

> Ürün Planlama ve Alokasyon › Alokasyon Uzmanı · Workers / Workless

Fazla stoklu ve tükenme riskli mağazaları eşleştirip **ürün-beden-renk bazında transfer önerisi** çıkarır. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Haftalık satış | Dönem satış adedi × 7 ÷ `--donem-gun` (varsayılan 28) |
| Kapsama (hafta) | Stok ÷ haftalık satış |
| Alıcı (tükenme riski) | Satışı var ve kapsama < `--min-hafta` (2). İhtiyaç = `--hedef-hafta` (4) × haftalık satış − stok, yukarı yuvarlanır |
| Verici (fazla stok) | Kapsama > `--max-hafta` (8) veya satışı yok. Fazla = stok − hedef hafta × haftalık satış, aşağı yuvarlanır; `--asgari-stok` adet mağazada bırakılır |
| Eşleştirme | Her ürün-beden-renk için alıcılar en düşük kapsamadan başlar; önce aynı bölgedeki, sonra fazlası en çok olan vericiden alınır |

Tam sınırdaki mağaza (kapsama tam 2 veya tam 8 hafta) dengeli sayılır. `--min-transfer` adedinden küçük transfer
önerilmez.

**Örnek:** 28 günde 16 satış, stok 2 → haftalık 4, kapsama 0,5 hafta → ihtiyaç 4 × 4 − 2 = 14. Aynı üründe 28 günde 8
satış, stok 40 → haftalık 2, kapsama 20 hafta → fazla 40 − 8 = 32. Öneri: 14 adet transfer.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --stok magaza_stok_satis.xlsx --donem-gun 28 --hedef-hafta 4 --min-hafta 2 --max-hafta 8 --asgari-stok 1
```

Sütunlar: Mağaza, Ürün, Stok, Satış (zorunlu); Bölge, Beden, Renk.

## Çıktı

`cikti/transfer_onerisi.xlsx`:

- `Transfer Önerileri`: gönderen, alan, ürün-beden-renk, adet; iki mağazanın stok ve kapsaması; boş "Onay" sütunu.
- `Karşılanamayan İhtiyaç`: transferle kapatılamayan eksikler (merkez depodan sevk veya sipariş gerekir).
- `Kalan Fazla Stok`: transferden sonra da fazla kalan kalemler (indirim veya iade adayı).
- `Mağaza Özeti`: mağaza başına gönderilecek ve alınacak adet.
- `Uyarılar`: okunamayan ve mükerrer satırlar.

## Dikkat

- **Eşleştirme sezgiseldir.** En iyi (en düşük maliyetli) dağıtımı garanti etmez; nakliye maliyeti, mesafe, koli içi
  adet ve mağaza depo kapasitesi hesaba katılmaz. Bölge yalnız "aynı bölge önce" tercihinde kullanılır.
- **Satışı olmayan ürün verici sayılır.** Yeni gelmiş veya rafa çıkmamış ürün de "satış yok" görünür; bu kalemleri
  göndermeden önce kontrol edin.
- **Stoksuz kalan mağazanın satışı düşük görünür.** Dönem içinde stok bitmişse haftalık satış gerçek talebin
  altındadır ve ihtiyaç eksik hesaplanır.
- **Beden dağılımı bozulabilir.** Kod her bedeni ayrı değerlendirir; vericide kırık seri kalıp kalmadığına bakmaz.
- **Eşikler varsayılandır.** 2, 4 ve 8 hafta genel bir ölçüt değildir; sezon sonuna yakınlık ve ürün grubuna göre
  ayarlayın.
- **Örnek veri:** Mağazalar, ürünler ve adetler kurgusaldır.

## Testler

- Alıcı ve verici sınırları, ihtiyaç ve fazla hesabı, asgari stok, farklı dönem uzunluğu.
- Eşleştirme sırası: aynı bölge önceliği, en acil alıcı, en küçük transfer, beden ayrımı.
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
