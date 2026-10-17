# Kesim Kat Planı · Kod Bloğu

> Tekstil ve Konfeksiyon › Planlama › Planlama Uzmanı · Gerisi Bende

Sipariş beden dağılımından **pastal (marker) oranlarını** ve **kat adetlerini** çıkarır. Örnek bir pastal:
S×1 · M×2 · L×2 · XL×1 × 80 kat. Beden tüketimi verilirse pastal boylarını ve kumaş ihtiyacını hesaplar.
İnternete bağlanmaz.

## Nasıl çalışır?

**Kısıtlar**
- **Azami kat** (`--max-kat`): Kumaş türüne ve kesim bıçağının yüksekliğine göre belirlenir.
- **Pastaldaki azami ürün** (`--max-urun`): Masa boyuna göre belirlenir.
- **Asgari kat** (`--min-kat`)
- **Hedef:** Sipariş × (1 + fazla kesim %).

**Yöntem**
1. **Pastal seçimi:** Her adımda kalan adetlerden **en çok ürünü kesen** pastal seçilir. Pastal oranı kalan
   adetlerle orantılı dağıtılır.
2. **Kapanış:** Kalan miktar küçüldüğünde, **fazla kesimi en az** olan tek ya da iki pastallık bir kapanış
   aranır. Kapanış pastalı da asgari kat kuralına uyar; 1-2 katlı pastal oluşmaz.
3. **Sonuç:** Her bedende hedef adet karşılanır. Fazla kesim en çok pastal kapasitesi kadar ya da hedefin %2'si
   kadar olur.

**Pastal boyu:** Σ (oran × beden tüketimi) + 2 × uç payı. Kumaş ihtiyacı = pastal boyu × kat.

> Yöntem sezgiseldir (greedy). Çoğu durumda iyi bir plan verir ama her zaman en az pastallı plan değildir.
> Farklı `--max-urun` ve `--max-kat` değerlerini deneyip karşılaştırın.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                           # örnek: 3 model-renk
python main.py --siparis siparis.xlsx --max-kat 80 --max-urun 6 --fazla-kesim 2
python main.py --siparis siparis.xlsx --max-kat 40 --max-urun 4 --tuketim tuketim.xlsx --uc-payi 3
```

| Dosya | Sütunlar |
|---|---|
| Sipariş | [Model,] Renk, S, M, L, XL … — veya Renk, Beden, Adet |
| Tüketim | Beden, Tüketim (m/adet) — pastal verimini içeren ortalama tüketim |

## Çıktı

`Kat Planı` (pastal no, oran, kat, kesilen adet, pastal boyu, kumaş) · `Sipariş ve Kesim` (beden bazında fark;
eksik kesim kırmızı) · `Bilgi`

## Dikkat

- **Pastal boyu tahmini:** Gerçek pastal boyu CAD pastal yerleşiminden çıkar. Buradaki boy ortalama tüketimle
  yapılan bir tahmindir.
- **Kumaş eni ve kusurlar:** Top boyu, renk tonu (şade) farkı ve kumaş kusurları kat planını değiştirebilir. Farklı
  şadeleri ayrı pastallarda kesin.

## Testler

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
