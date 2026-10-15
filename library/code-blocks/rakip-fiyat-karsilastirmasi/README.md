# Rakip Fiyat Karşılaştırması · Kod Bloğu

> Satın Alma ve Kategori Yönetimi › Kategori Uzmanı · Workers / Workless

Rakip fiyat taramalarını **ürün koduna ya da barkoda göre** bizim fiyat listesiyle eşleştirir; ürün, kategori ve
rakip bazında **fiyat endeksi** çıkarır. İnternete bağlanmaz; fiyatları siz toplarsınız.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Eşleştirme | Taramadaki kod bizim ürün koduna ya da barkoda eşitse |
| Geçerli fiyat | Her ürün-rakip çiftinin en son kampanyasız taraması |
| Eski tarama | En son tarama tarihinden `--eski` (varsayılan 30) günden eski → endekse girmez |
| Şüpheli eşleşme | Rakip fiyatı bizimkinin yarısından az ya da iki katından fazla → endekse girmez |
| Ürün endeksi | Bizim fiyat ÷ rakiplerin ortalama fiyatı × 100 |
| Sepet endeksi | Σ(bizim fiyat × satış adedi) ÷ Σ(rakip fiyatı × satış adedi) × 100 |

Endeks **100'ün üstündeyse rakipten pahalıyız**. Sepet endeksi rakip ve kategori bazında, yalnız o rakipte geçerli
fiyatı olan ürünlerle hesaplanır; satış adedi yazılmamışsa her ürün eşit ağırlık alır.

**Uyarılar:** ürün endeksi 100 + `--esik` (varsayılan 5) üstünde → "Rakip ortalamasından pahalı"; 100 − `--esik`
altında → "Rakip ortalamasından ucuz" (marj fırsatı olabilir).

**Örnek:** Pirinç bizde 74,90; Rakip A'da 69,90 (C'deki tarama 30 günden eski) → endeks 74,90 ÷ 69,90 × 100 =
**107,15**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                  # örnek veri
python main.py --urunler bizim_fiyatlar.xlsx --taramalar rakip_taramalari.xlsx
python main.py --urunler bizim_fiyatlar.xlsx --taramalar rakip_taramalari.xlsx --esik 3 --eski 14
```

Zorunlu sütunlar: bizim fiyatlarda Ürün Kodu ve Fiyat; taramalarda Rakip, Ürün Kodu (ya da Barkod) ve Rakip Fiyat.
Kampanya sütununa "Evet" ya da "x" yazılan taramalar kampanyalı sayılır.

## Çıktı

`cikti/rakip_fiyat_karsilastirmasi.xlsx`:

- `Endeks Özeti`: rakip × kategori ağırlıklı endeks ve rakip toplamı (TÜMÜ).
- `Ürün Karşılaştırma`: ürün başına her rakibin fiyatı (bizden ucuz olanlar sarı), ortalama, en düşük, iki endeks.
- `Taramalar`: endekse giren, kampanyalı, eski ve şüpheli taramalar.
- `Uyarılar`: pahalı / ucuz ürünler, şüpheli eşleşmeler, eşleşmeyen tarama kodları.

## Dikkat

- **Eşleştirme sizin sorumluluğunuzdadır.** Farklı gramaj, ambalaj ya da paket adedindeki ürünler aynı kodla
  eşleştirilirse endeks yanlış çıkar. Kod yalnız yarı / iki kat sınırıyla kaba bir kontrol yapar; birim fiyata
  (kg, L) çevirme yapmaz.
- **Fiyat toplama yöntemi:** Rakip fiyatlarını mağaza ziyareti ya da rakibin kendi kanallarından toplarken rakibin
  internet sitesi kullanım koşullarına ve rekabet hukukuna uyun; rakiplerle fiyat bilgisi paylaşmak ya da fiyatları
  birlikte belirlemek rekabet hukuku açısından risklidir. Bu paket yalnız sizin topladığınız veriyi karşılaştırır.
- **Eşik politika varsayılanıdır.** %5 bant ve 30 günlük geçerlilik genel kabul görmüş değerler değildir; kategori
  stratejinize göre değiştirin.
- **Raf fiyatı ile ödenen fiyat farklı olabilir.** Kart, sadakat programı ve çoklu alım indirimleri endekse girmez.
- **Örnek veri:** Ürünler, barkodlar, rakip adları ve fiyatlar kurgusaldır.

## Testler

- En son kampanyasız fiyatın seçimi, barkodla eşleşme, kampanyalı / eski / şüpheli taramaların ayrılması.
- Ürün endeksi, en ucuz rakip, ağırlıklı kategori ve rakip endeksi; eşikler, okuma ve komut satırı.

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
