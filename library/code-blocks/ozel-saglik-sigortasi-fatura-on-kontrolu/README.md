# Özel Sağlık Sigortası Fatura Ön Kontrolü

> Anlaşmalı Kurumlar ve Faturalama › Anlaşmalı Kurumlar Uzmanı · Gerisi Bende

Sigorta şirketine gönderilecek fatura kalemlerini **anlaşma fiyatları, provizyon onayı ve poliçe bilgileriyle**
karşılaştırır, **kesinti riski taşıyan kalemleri ve riskli tutarı** nedenleriyle listeler. Yapay zekâ kullanmaz;
internete bağlanmaz.

## Kontroller

**Kalem bazında**

| Kontrol | Riskli tutar |
|---|---|
| Hesap: adet × birim fiyat ≠ tutar | fark |
| Hizmet kodu anlaşma listesinde yok | tutarın tamamı |
| Birim fiyat > anlaşma fiyatı | (birim − anlaşma) × adet |
| Provizyon gerekli ama yok, başka dosyaya ait, hizmet tarihi onay tarihi + geçerlilik günü dışında, kod onaylı kodlarda yok | tutarın tamamı |
| Hizmet tarihi poliçe süresi dışında ya da kod poliçede kapsam dışı | tutarın tamamı |
| Dosyada aynı kodun toplam adedi > adet sınırı | aşan adet × birim fiyat |
| Mükerrer kalem (aynı dosya, kod, tarih, fiyat) | tutarın tamamı |

Bir kalemde birden fazla neden varsa riskli tutar en yüksek nedene göre alınır ve kalem tutarını aşmaz.

**Dosya bazında:** dosya toplamı > provizyon onaylı tutar; şirket payı (toplam × (1 − katılım payı)) > kalan poliçe
limiti (yıllık limit − kullanılan).

## Kullanım

```bash
python main.py                                         # örnek (kurgusal) 4 dosya, 10 kalem
python main.py --fatura ekim_faturalari.xlsx --anlasma sirket_x_fiyatlar.xlsx --provizyon provizyonlar.xlsx --police policeler.xlsx
```

## Çıktı

- `cikti/oss_fatura_kontrol.xlsx`: Kalemler (kırmızı: riskli tutar var), Dosyalar, Özet (toplam, riskli tutar,
  neden bazında kalem sayısı).

## Dikkat

- **Ön kontroldür.** Sigorta şirketleri provizyon, ön onay, tıbbi gereklilik ve fatura inceleme kurallarını kendi
  sözleşme ve sistemlerine göre uygular. Anlaşma fiyat listesi, provizyon geçerlilik süresi ve poliçe kapsamı şirketten
  şirkete değişir; listelerin güncel olduğunu kontrol edin.
- **Kişisel veri.** Tabloda hasta adı ve kimlik bilgisi gerekmez; dosya ve poliçe numarası yeterlidir.
- **Örnek veri:** Hizmet kodları, fiyatlar, poliçeler ve provizyonlar kurgusaldır.

## Testler

Testler; her kalem kuralını (fiyat farkı, provizyon yok / süresi geçmiş / başka dosya / onaysız kod, poliçe süresi ve
kapsam dışı, anlaşma dışı kod, adet sınırı, hesap hatası, mükerrer kalem), dosya kontrollerini (provizyon tutarı, katılım
payı, kalan limit) ve örnek veride toplam 21.630,00 TL riskli tutarı doğrular.

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
