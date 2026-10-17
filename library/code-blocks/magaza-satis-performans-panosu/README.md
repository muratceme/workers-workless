# Mağaza Satış Performans Panosu · Kod Bloğu

> Mağaza Operasyonları › Mağaza Müdürü · Gerisi Bende

Günlük satış verisinden **ciro, fiş adedi, sepet ortalaması, dönüşüm oranı ve geçen yıl karşılaştırmasını (LFL)**
mağaza ve gün bazında tek panoda toplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Sepet ortalaması | Ciro ÷ fiş adedi |
| Dönüşüm oranı % | Fiş adedi ÷ ziyaretçi sayısı × 100 |
| Fiş başına adet | Satılan adet ÷ fiş adedi |
| LFL büyüme % | (Ciro − geçen yıl ciro) ÷ geçen yıl ciro × 100 |
| Hedef gerçekleşme % | Ciro ÷ hedef ciro × 100 |
| Ciro payı % | Mağazanın cirosu ÷ toplam ciro × 100 |

Her oran yalnız ilgili verisi yazılmış günler üzerinden hesaplanır: dönüşümde ziyaretçi sayısı, LFL'de geçen yıl
cirosu, hedef gerçekleşmede hedef cirosu olan günler. Mağaza ve toplam oranları günlük oranların ortalaması değil,
**toplamların oranıdır**.

Uyarılar: mağazanın LFL büyümesi `--lfl` (−5) yüzdesinin altındaysa; hedef gerçekleşmesi `--hedef` (90) yüzdesinin
altındaysa; geçen yıl verisi hiç yoksa (bilgi).

**Örnek:** Haftalık ciro 365.000, fiş 2.400 → sepet ortalaması **152,08**. Ziyaretçi 12.700 → dönüşüm **%18,90**.
Geçen yılın aynı günlerinde ciro 394.000 → LFL **−%7,36**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                             # örnek veri
python main.py --satislar gunluk_satislar.xlsx
python main.py --satislar gunluk_satislar.xlsx --lfl 0 --hedef 95
```

Zorunlu sütunlar: Tarih, Mağaza, Ciro. Her satır bir mağazanın bir günüdür.

## Çıktı

`cikti/magaza_satis_panosu.xlsx`:

- `Mağazalar`: ciroya göre sıralı; bütün göstergeler ve ciro payı; son satır toplam.
- `Günler`: bütün mağazaların günlük toplamı, haftanın günüyle.
- `Günlük Pano`: mağaza × gün ayrıntısı.
- `Uyarılar`.

## Dikkat

- **Geçen yılın hangi günle eşleşeceğini siz belirlersiniz.** Aynı takvim tarihi yerine aynı haftanın aynı gününü
  (ör. geçen yılın karşılık gelen cumartesisi) yazmak perakendede daha anlamlıdır; bayram, kampanya ve tatil
  kaymalarını da siz gözetirsiniz. Kod yalnız yazdığınız rakamı karşılaştırır.
- **LFL yalnız karşılaştırılabilir günleri içerir.** Geçen yıl cirosu boş ya da sıfır olan günler (yeni açılan
  mağaza, tadilat) LFL'ye girmez; bu yüzden LFL'deki ciro toplam cirodan küçük olabilir.
- **Nominal karşılaştırmadır.** LFL büyüme enflasyondan arındırılmaz; fiyat artışı dönemlerinde adet ve fiş
  sayısındaki değişime de bakın.
- **Dönüşüm, ziyaretçi sayacına bağlıdır.** Sayaç arızası, personel giriş-çıkışı ve grup hâlinde gelen müşteriler
  oranı bozar; mağazalar arasında sayaç türü farklıysa karşılaştırma yanıltır.
- **Ciro tanımı sizindir.** KDV dahil mi hariç mi, iadeler düşülmüş mü; ciro, hedef ve geçen yıl aynı tanımda olmalı.
- **Kapsam dışı:** metrekare verimliliği, brüt kâr, stok devir hızı, personel verimliliği, saatlik dağılım.
- **Örnek veri:** Mağazalar ve rakamlar kurgusaldır.

## Testler

- Sepet ortalaması, dönüşüm, fiş başına adet, LFL ve hedef gerçekleşme; verisi eksik günlerin oranlara girmemesi.
- Mağaza, gün ve toplam düzeyinde toplamların oranı; ciro payı.
- LFL ve hedef eşikleri, geçen yıl verisi olmayan mağaza, okuma ve komut satırı.

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
