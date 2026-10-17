# Fiyat Listesi Güncelleme · Kod Bloğu

> Satış › Satış Destek Uzmanı · Gerisi Bende

Maliyeti değişen ürünler için **yeni fiyatı sizin kurallarınıza göre** hesaplar: hedef marj, yuvarlama basamağı,
artış tavanı, asgari fiyat. Yeni fiyat listesini ve her ürün için "neden bu fiyat" açıklamasını verir. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Brüt marj % | (Fiyat − maliyet) ÷ fiyat × 100 |
| Hedef | Ürün satırındaki hedef → yoksa `--marj` → o da yoksa ürünün mevcut marjı korunur |
| Ham fiyat | `--yontem marj`: yeni maliyet ÷ (1 − hedef ÷ 100) · `--yontem katki`: yeni maliyet × (1 + hedef ÷ 100) |
| Yuvarlama | Ham fiyat `--yuvarla` basamağının katına **yukarı** yuvarlanır |
| İndirim yapmama | Hesaplanan fiyat mevcut fiyatın altındaysa mevcut fiyat korunur (`--indirim` verilmedikçe) |
| Artış tavanı | Artış `--azami-artis` yüzdesini aşarsa fiyat mevcut fiyat × (1 + tavan) ile sınırlanır |
| Asgari fiyat | Satırdaki asgari fiyatın altına inilmez (tavandan sonra uygulanır) |
| Asgari değişim | Değişim `--asgari-degisim` yüzdesinin altındaysa fiyat değiştirilmez |

**Örnek:** Mevcut fiyat 250, yeni maliyet 195, hedef marj %40 → ham fiyat 195 ÷ 0,60 = **325,00** (%30 artış).
Tavan %20 → **300,00**; bu fiyatla marj %35 kalır ve rapor bunu uyarı olarak yazar.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri: yuvarlama 0,50, artış tavanı %20
python main.py --urunler urunler.xlsx
python main.py --urunler urunler.xlsx --marj 35 --yuvarla 0,50 --azami-artis 15 --asgari-degisim 2
python main.py --urunler urunler.xlsx --yontem katki --marj 40 --indirim
```

Zorunlu sütunlar: Ürün Kodu, Mevcut Fiyat. Yeni maliyeti boş olan ürünün fiyatı değiştirilmez.

## Çıktı

`cikti/fiyat_listesi_guncelleme.xlsx`:

- `Yeni Fiyat Listesi`: ürün kodu, ürün, kategori, yeni fiyat (yayımlamaya hazır sade liste).
- `Değişim Raporu`: eski marj, fiyat değişmezse marj, hedef, yeni fiyat, değişim, yeni marj, uygulanan kural.
- `Özet`: kategori bazında artan, inen, aynı kalan ürün sayısı, ortalama değişim ve ciro etkisi.
- `Uyarılar`: maliyetin altında kalan fiyat, tavan yüzünden hedef marjı tutmayan ürün, eksik maliyet.

## Dikkat

- **Fiyat kararı sizindir.** Kod maliyet ve marj kuralını uygular; rakip fiyatına, talebe, müşteri sözleşmesine,
  bayi anlaşmasına ya da kampanyaya bakmaz. Listeyi yayımlamadan önce gözden geçirin.
- **Marj ile maliyet üstü kâr aynı şey değildir.** Maliyeti 60, fiyatı 100 olan üründe marj %40, maliyet üstü kâr
  %66,67'dir. Hedef sütununu hangi anlamda doldurduysanız `--yontem` onu söylemelidir.
- **Fiyatlar KDV hariç ve tek para biriminde olmalı.** Kod KDV eklemez, kur çevirmez.
- **Ciro etkisi bir tahmindir.** Satış adedinin fiyat değişse de aynı kalacağını varsayar.
- **Mevzuat:** Etiket, tarife ve fiyat listesi duyurusu, haksız fiyat artışı ve sözleşmeli müşteriye bildirim
  süresi gibi yükümlülükler bu paketin kapsamında değildir; sektörünüzdeki kuralları ayrıca kontrol edin.
- **Kapsam dışı:** kademeli (miktara göre) fiyat, müşteri ve bayi iskonto yapısı, dövizli liste, psikolojik fiyat
  sonları (9,90 gibi), rakip fiyat karşılaştırması.
- **Örnek veri:** Ürünler, fiyatlar ve maliyetler kurgusaldır.

## Testler

- Mevcut marjı koruma, satır ve komut satırı hedefi, maliyet üstü kâr yöntemi, yukarı yuvarlama basamağı.
- İndirim yapmama ve `--indirim`, artış tavanı sınırı (tam tavanda ve üstünde), asgari fiyat, asgari değişim.
- Maliyetin altında fiyat, hedef marjın altında kalma, eksik maliyet, kategori özeti ve ciro etkisi, okuma, komut satırı.

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
