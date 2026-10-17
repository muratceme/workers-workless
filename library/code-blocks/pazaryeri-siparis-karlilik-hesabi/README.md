# Pazaryeri Sipariş Kârlılık Hesabı · Kod Bloğu

> E-ticaret › Pazaryeri Yönetimi › Pazaryeri Uzmanı · Gerisi Bende

Pazaryeri sipariş dökümünü ürün maliyet listesiyle birleştirip **sipariş satırı ve ürün bazında net kârı**
hesaplar. Zararına satılan ürünleri ve **başabaş satış fiyatını** gösterir. Pazaryerinin size ödeyeceği tutarı
(e-ticaret stopajı dahil) ve satışın KDV etkisini ayrıca verir. İnternete bağlanmaz.

## Hesap

```
Net satış     = Satış tutarı (KDV dahil, satıcı indirimi düşülmüş) ÷ (1 + ürün KDV oranı)
Net kâr       = Net satış − komisyon − kargo − hizmet bedeli (hepsi KDV hariç) − ürün maliyeti (KDV hariç)
Pazaryeri ödemesi = Satış tutarı − komisyon − kargo − hizmet bedeli (KDV dahil) − stopaj
```

- **Komisyon:** `komisyon_kdv_dahil: true` ise KDV dahil satış tutarı × oran olarak hesaplanır ve bu tutar KDV'yi
  içerir. Trendyol Satıcı Bilgi Merkezi'ne göre komisyon faturaları "KDV dahil rakamlardan hesaplanarak"
  kesilir, ekstra KDV ödenmez. Bazı hesaplayıcılar komisyona ayrıca %20 KDV ekliyor; bu, komisyonu fazla
  gösterir. `false` ise KDV hariç tutar × oran hesaplanır ve üzerine KDV eklenir. Sipariş dosyasında Komisyon
  sütunu varsa gerçek tutar kullanılır.
- **Kargo:** Sipariş dosyasındaki tutar, yoksa ürün desilerinin toplamına göre desi tarifesi, o da yoksa sabit
  bedel. Çok kalemli siparişte kargo ve hizmet bedeli satırlara satış tutarı oranında dağıtılır.
- **İade:** Satış, komisyon ve ürün maliyeti geri alınır; gidiş ve dönüş kargosu ile hizmet bedeli kayıp yazılır.
  İptaller hesaba girmez.
- **E-ticaret stopajı:** Aracı hizmet sağlayıcı, KDV hariç satış tutarı üzerinden %1 keser (9284 sayılı
  Cumhurbaşkanı Kararı, 01.01.2025). Gelir/kurumlar vergisinden mahsup edildiği için kârı değil nakit akışını
  etkiler.
- **Başabaş fiyat:** Tek adetlik siparişte kârı sıfırlayan KDV dahil satış fiyatıdır. Ortalama satış fiyatı bunun
  altındaysa ürün işaretlenir.

> `ornek_veri/kanallar.json` içindeki komisyon, desi ve hizmet bedelleri **örnektir**. Bu değerler sık
> değişir; satıcı panelinizdeki güncel değerlerle değiştirin.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                   # örnek siparişlerle dener
python main.py --siparisler siparisler.xlsx --maliyetler urun_maliyetleri.xlsx --kanallar kanallar.json
python main.py --siparisler trendyol.xlsx --maliyetler maliyet.xlsx --kanallar kanallar.json --kanal Trendyol
```

## Çıktı

`Kanal Özeti` · `Ürün Kârlılığı` (en zararlıdan sıralı) · `Sipariş Detayı` · `Bilgi`

## Testler

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| E-ticaret ödemelerinde %1 stopaj — 9284 sayılı Cumhurbaşkanı Kararı (RG 22.12.2024; yürürlük 01.01.2025) | [Erdem & Erdem](https://www.erdem-erdem.av.tr/bilgi-bankasi/elektronik-ticaret-uzerinden-yapilan-odemeler-icin-stopaj-uygulamasi), [T24](https://www.t24.com.tr/index.php/yazarlar/erdogan-saglam/9-soruda-e-ticarete-getirilen-yuzde-1-stopaj-uygulamasi,48076) | Teyit edildi (ikincil kaynaklar) |

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
