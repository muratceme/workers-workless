# Hakediş Hesaplama · Kod Bloğu

> İnşaat › Teknik Ofis › Teknik Ofis Mühendisi · Gerisi Bende

Birim fiyatlı sözleşmenin **poz listesi** ve dönem **metrajlarından** hakediş icmalini hazırlar:
kümülatif metraj özeti, *toplam − önceki = bu hakediş* tablosu, KDV, KDV tevkifatı, stopaj, teminat ve
avans kesintileri, ödenecek net tutar ve tüm hakedişlerin geçmişi. İnternete bağlanmaz.

## Vergi ve kesinti kuralları

| Kalem | Varsayılan | Dayanak |
|---|---|---|
| KDV | %20 | `kdv_orani` |
| KDV tevkifatı | **4/10**; alıcı belirlenmiş alıcıysa ya da KDV dahil sözleşme bedeli ≥ 5.000.000 TL ise (`otomatik`) | KDV Genel Uygulama Tebliği I/C-2.1.3.2.1 (Seri No: 35 ile 01.03.2021'den beri) |
| Stopaj | **%5**, yıllara yaygın inşaat ve onarım işlerinde (`yillara_yaygin: true`) | GVK md. 94/3, KVK md. 15; 3491 sayılı CBK. Demiryolu, gemi ve nükleer santral inşaatında %1 |
| Damga vergisi | 0 (sözleşmeniz hakedişten kesilmesini öngörüyorsa oranı girin) | `damga_vergisi_orani` |
| Teminat kesintisi | 0 | `teminat_kesintisi_orani` |
| Avans mahsubu | Bu hakediş × avans / sözleşme bedeli; toplamı avansı aşmaz | `avans_tutari` |
| Fiyat farkı | Hesaplanmış tutar girilir (bu paket fiyat farkı formülünü uygulamaz) | `fiyat_farki` |

Tutarlar **kümülatif** hesaplanır, bu hakediş önceki hakedişlerden farkla bulunur; yuvarlama farkı birikmez.

## Kontroller

- Kümülatif miktarı **sözleşme miktarını aşan pozlar** (iş artışı / onaylı fazla metraj gerekir)
- Toplam iş tutarı sözleşme bedelini aşarsa iş artışı uyarısı (kamu işlerinde 4735 sayılı Kanun md. 24)
- Metrajda olup **poz listesinde olmayan** pozlar (yeni birim fiyat gerekir) hesaba katılmaz ve listelenir

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                   # örnek sözleşmeyle 3 no'lu hakediş
python main.py --sozlesme sozlesme.json --pozlar pozlar.xlsx --metraj metraj.xlsx --hakedis 4
```

`ornek_veri/sozlesme.json` dosyasını kendi sözleşmenize göre kopyalayıp düzenleyin.

## Çıktı

`Hakediş İcmali` · `Metraj Özeti` · `Hakediş Geçmişi` · `Bilgi`

> Sonuçlar sözleşme şartlarına ve güncel mevzuata göre mali müşavir ve kontrol mühendisi tarafından
> teyit edilmelidir. Oranlar değişirse `sozlesme.json` üzerinden güncelleyin.

## Testler

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 193 sayılı Gelir Vergisi Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.4.193.pdf) | Resmî metinle karşılaştırılan maddeler: 25, 94 |
| 5520 sayılı Kurumlar Vergisi Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.5520.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |
| 4735 sayılı Kamu İhale Sözleşmeleri Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4735.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |
| Yapım işlerinde 4/10 KDV tevkifatı; KDV mükellefleri için KDV dahil 5 milyon TL eşiği — 35 Seri No'lu KDV GUT Değişikliği (RG 16.02.2021 / 31397) | [TÜRMOB (Resmî Gazete metni)](https://www.turmob.org.tr/arsiv/mbs/resmigazete/31397-31397-1-kdv.pdf) | Teyit edildi (Resmî Gazete metni) |
| Yıllara yaygın inşaat ve onarım işlerinde %5 stopaj (GVK md. 94/3, 3491 sayılı CBK) | — | Teyit edilemedi: aramayla doğrulanmadı; güncel oranı mali müşavirinizle kontrol edin |

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
