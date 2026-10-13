# Depo Yerleşim (ABC) Önerisi · Kod Bloğu

> Lojistik ve Depo › Depo Sorumlusu · Workers / Workless

Ürünleri **rafa kaç kez gidildiğine** göre ABC sınıflarına ayırır ve sık çıkanların sevkiyata yakın bölgede
durmasını önerir. Mevcut bölgeyi verirseniz hangi ürünün öne, hangisinin geriye alınacağını listeler. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Çıkış sıklığı | Ürünün çıkış (toplama) satırı sayısı; miktar değil |
| Pay %, kümülatif % | Sıklık ÷ toplam sıklık × 100; sıklığa göre azalan sırada birikimli pay |
| A sınıfı | Kendinden önceki kümülatif pay `--a` (%80) altında olan ürünler |
| B sınıfı | Kendinden önceki kümülatif pay `--b` (%95) altında olanlar |
| C sınıfı | Kalan ürünler |
| D sınıfı | Ürün listesinde olup dönemde hiç çıkışı olmayanlar |
| Önerilen bölge | A sınıfı → A bölgesi (sevkiyata en yakın), B → B, C ve D → C |
| Öneri | Öne al · Geriye al · Yerinde (mevcut bölge önerilenle karşılaştırılır) |

**Örnek:** Toplam 1.000 çıkışın 320'si tek üründe (%32), ilk dört ürün toplamın %80'ini oluşturuyor → dördü de
**A**. Beşinci ürün için önceki kümülatif pay tam %80 olduğundan **B** olur.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri
python main.py --cikislar cikislar.xlsx --urunler urunler.xlsx
python main.py --cikislar cikislar.xlsx --a 70 --b 90     # ürün listesi olmadan yalnız sınıflandırma
```

Çıkış dosyası iki biçimde olabilir: her satır bir toplama satırı (Tarih, Stok Kodu, Miktar) ya da özet (Stok Kodu,
Çıkış Sayısı, Miktar). "Çıkış Sayısı" sütunu yoksa her satır bir çıkış sayılır. Zorunlu sütun: Stok Kodu.

## Çıktı

`cikti/depo_yerlesim_abc.xlsx`:

- `Sınıflandırma`: sıklık, miktar, pay, kümülatif pay, sınıf, mevcut ve önerilen bölge, öneri.
- `Taşıma Listesi`: önce öne alınacaklar (en sık çıkandan), sonra geriye alınacaklar; yeni konumu siz yazarsınız.
- `Özet`: sınıfların ürün ve çıkış payı; mevcut bölge × sınıf matrisi.
- `Uyarılar`: C bölgesindeki A sınıfı ürünler, A bölgesinde yer tutan çıkışsız ürünler.

## Dikkat

- **Bölgeleri siz tanımlarsınız.** A bölgesi sevkiyat alanına en yakın ve erişimi en kolay raflar, C en uzak
  olanlardır; kod depo krokinizi bilmez, ürün dosyasına yazdığınız bölgeyi kullanır.
- **Öneri yalnız sıklığa bakar.** Ürünün boyutu, ağırlığı, istif ve raf kapasitesi, soğuk zincir, tehlikeli madde
  ayrımı, birlikte sipariş edilen ürünler ve A bölgesinde yeterli yer olup olmadığı hesaba katılmaz. Taşıma
  listesini bu kısıtlarla birlikte değerlendirin.
- **Dönem seçimi sonucu değiştirir.** Mevsimlik ürünlerde tek aylık veri yanıltır; en az birkaç aylık çıkış
  kullanın ve yerleşimi belirli aralıklarla yenileyin.
- **Sıklık miktardan farklıdır.** Ayda bir kez 1.000 adet çıkan ürün C, her gün birer adet çıkan ürün A olabilir;
  yerleşim için doğru ölçü rafa gidiş sayısıdır. Tüketim değerine göre ABC için Stok ABC-XYZ Analizi paketine bakın.
- **Taşımanın da maliyeti vardır.** Sınırdaki ürünler dönemden döneme sınıf değiştirebilir; her çalıştırmada
  bütün listeyi taşımak yerine farkı büyük olanlardan başlayın.
- **Örnek veri:** Ürünler, konumlar ve çıkış sayıları kurgusaldır.

## Testler

- Sınıf sınırları (önceki kümülatif pay tam eşikte), eşit sıklıkta sıralama, özel `--a` ve `--b` değerleri.
- Çıkışsız ürün (D), ürün listesinde olmayan kod, bölgesi boş ürün, öne ve geriye alma, taşıma listesi sırası.
- Satır bazlı ve özet çıkış dosyası, sınıf özeti ve bölge matrisi, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
