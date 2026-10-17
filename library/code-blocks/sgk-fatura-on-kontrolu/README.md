# SGK Fatura Ön Kontrolü

> Anlaşmalı Kurumlar ve Faturalama › SGK Faturalama Uzmanı · Gerisi Bende

Medula'ya gönderilecek hizmet kalemlerini **kurumunuzun tanımladığı SUT kurallarına** göre ön kontrolden geçirir:
tanı-işlem uyumu, adet sınırı, tekrar süresi, branş şartı ve birlikte faturalanamayan işlemler. Kurallar bir tabloda
durur; kodun içinde SUT maddesi yoktur. Yapay zekâ kullanmaz; internete bağlanmaz.

## Kural türleri

`kurallar.csv` sütunları: İşlem Kodu, Kural, Değer, Kapsam, Açıklama.

| Kural | Değer | Kapsam | İhlal |
|---|---|---|---|
| Tanı Gerekli | ICD-10 önekleri, ör. `I20 I25 R07` | — | Kalemin tanılarından hiçbiri öneklerden biriyle başlamıyor |
| Adet Sınırı | sayı | Takip (varsayılan) / Gün (hasta + gün) / Yıl (hasta + takvim yılı) | Toplam adet sınırı aşıyor (aşan adet × birim tutar risk) |
| Tekrar Süresi | gün | — | Aynı hasta, aynı işlem, önceki işlemden bu kadar gün geçmeden |
| Branş Şartı | branşlar, `;` ile | — | Kalemin branşı listede yok |
| Birlikte Faturalanamaz | işlem kod(lar)ı | Takip (varsayılan) / Gün | Aynı takipte (ya da aynı hasta ve günde) birlikte var |

**Genel kontroller (her kalem):** tanı kodu yok, ICD-10 biçimi geçersiz (ör. `I209`), adet 0 ya da negatif, mükerrer
kalem (aynı takip, işlem ve tarih).

## Kullanım

```bash
python main.py                                                     # örnek (kurgusal) 9 kalem, 6 kural
python main.py --hizmetler ekim_hizmetleri.xlsx --kurallar sut_kurallarimiz.xlsx
```

Tekrar süresi kontrolünün önceki dönemleri de görebilmesi için hizmet dosyasına ilgili hastaların önceki işlemlerini de
ekleyin.

## Çıktı

- `cikti/sgk_fatura_kontrol.xlsx`: Kalemler (kırmızı: ihlal var), Kurallar (kural bazında ihlal sayısı ve toplam).

## Dikkat

- **Kurallar sizin.** Örnek `kurallar.csv` ve işlem kodları (`ÖRN-…`) **kurgusaldır**, gerçek SUT kodu ya da kuralı
  değildir. Kuralları güncel Sağlık Uygulama Tebliği (SUT) ve eklerinden, kurumunuzun faturalama birimi tanımlamalıdır.
  SUT sık değişir; kural tablosuna değişiklik tarihini ve dayanağını "Açıklama" sütununda yazın.
- **Ön kontroldür.** Medula'nın kendi kontrolleri ve SGK'nın fatura inceleme sonuçları ayrıca uygulanır.
- **Kişisel veri.** Hasta adı ve TC kimlik no gerekmez; kurum içi hasta numarası yeterlidir.

## Testler

Testler; her kural türünü (tanı öneki, branş listesi, adet sınırı takip / gün / yıl kapsamı ve aşan adet tutarı, tekrar
süresi sıralaması, birlikte faturalanamaz takip / gün kapsamı), genel kontrolleri, tanınmayan kural ve sayısal olmayan
değer hatalarını ve örnek veride 2.420,00 TL riskli tutarı doğrular.

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
