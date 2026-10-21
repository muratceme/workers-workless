# Uygulama Abonelik Gelir Raporu · Kod Bloğu

> Sizden Gelenler › Abonelikli mobil uygulama · Gerisi Bende

Google Play **kazanç raporu** ile App Store Connect **özet satış raporunu** tek tabloda birleştirir: satış, iade,
mağaza komisyonu ve **net gelir**; günlük, aylık, ürün ve ülke kırılımında. İsteğe bağlı abonelik olayları dosyasından
**iptalleri** sayar. İnternete bağlanmaz; mağaza hesabınıza erişmez, yalnız indirdiğiniz dosyaları okur.

## Nasıl hesaplar?

| Platform | Kaynak sütun | Hesap |
|---|---|---|
| Google Play | `Transaction Type` = `Charge` | Satış (brüt) |
| Google Play | `Charge refund` | İade |
| Google Play | `Google fee` + `Google fee refund` | Mağaza komisyonu |
| Google Play | diğer türler | "Diğer" olarak nete katılır, adı korunur ve uyarı verilir |
| Google Play | `Amount (Merchant Currency)` | Tutar (satıcı para birimi) · Net = günün bütün satırlarının toplamı |
| App Store | `Units` × `Developer Proceeds` | Net (geliştirici payı; vergi ve Apple komisyonu düşülmüş). İadede `Units` negatiftir |
| App Store | `Subscription` = `New` / `Renewal` | Yeni abonelik / yenileme adedi |

- **Efektif komisyon oranı (Play)** = komisyon ÷ brüt satış.
- **İade oranı** = iade adedi ÷ satış adedi.
- **App Store'da brüt ve komisyon hesaplanmaz:** Apple'ın özet satış raporunda müşteri fiyatı vergi dahildir ve
  komisyon ayrı bir sütun değildir; uydurmak yerine yalnız net gösterilir.
- **Para birimi:** Raporlama para birimi `--para-birimi` (varsayılan Play satıcı para birimi). Diğer para birimleri
  `--kur` dosyasıyla çevrilir; kuru verilmeyen para birimi toplamlara **katılmaz** ve "Yüksek" uyarı verilir.

## Dosyaları nereden alırsınız?

- **Google Play:** Play Console'da raporları indirme bölümünden aylık kazanç raporu (`earnings_YYYYMM.zip`). ZIP'i
  açmadan verebilirsiniz; Google bu raporları UTF-16 kodlamayla verir, paket bunu kendisi tanır.
- **App Store:** App Store Connect'ten özet satış raporu (günlük ya da aylık `.txt`, sıkıştırılmışsa `.gz`).
- **İptaller:** Kazanç ve satış raporlarında iptal yoktur. Mağazanın abonelik raporlarından ya da kendi sisteminizden
  `Tarih; Olay; Adet` (isteğe bağlı `Platform; Ürün`) sütunlarıyla bir dosya hazırlayın. Olay adında "iptal / cancel"
  geçen satırlar iptal sayılır.

Birden çok ay için dosyaları ya da klasörü verin: `--play earnings/ --apple satis/`.

## Kullanım

```bash
python main.py                                    # örnek verilerle dener (Eylül 2026, kurgusal)
python main.py --play earnings_202609.zip --apple satis_202609.txt --kur kur.csv --olaylar abonelik_olaylari.csv
python main.py --play earnings/ --apple satis/ --para-birimi TRY --kur kur.csv
```

Terminal kullanmıyorsanız klasördeki `BASLAT.bat` (Windows) ya da `BASLAT.command` (macOS) dosyasına çift tıklayın.

## Çıktı

`cikti/abonelik_gelir_raporu.xlsx`: Özet, Günlük, Aylık, Ürünler, Ülkeler, Abonelik Olayları, İşlemler (her satır,
çevrilmiş tutarıyla; kuru olmayanlar kırmızı), Uyarılar.

`cikti/uygulama-abonelik-gelir-raporu.ozet.json`: günlük net, brüt, iade, komisyon ve iptal sayıları. **Uygulama
Yönetim Özeti** paketi bu dosyayı okuyup günlük / haftalık / aylık raporu hazırlar.

## Sınırlar

- Rapor, mağazaların **indirdiğiniz dosyadaki** rakamlarını toplar; ödeme (payout) mutabakatı yapmaz.
- Play Console ve App Store Connect sütun adlarını değiştirebilir. Paket sütunları adlarıyla eşleştirir (sırasına
  güvenmez); bir sütun bulunamazsa hangisinin eksik olduğunu söyler.
- Örnek kurlar ve tutarlar kurgusaldır.

## Kaynaklar ve dayanaklar

Kontrol tarihi: 21.10.2026. Sütun adları ve işlem türleri aşağıdaki belgelerden alınmıştır; platformlar biçimi
değiştirebilir, kendi dosyanızın başlık satırıyla karşılaştırın.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| Google Play Console: rapor türleri (kazanç ve tahmini satış raporu sütunları, UTF-16) | [support.google.com](https://support.google.com/googleplay/android-developer/answer/6135870) | Sütun adları belgeyle karşılaştırıldı |
| Google Play Console: satış ve ödeme raporlarını indirme (Charge, Charge refund, Google fee, Google fee refund) | [support.google.com](https://support.google.com/googleplay/android-developer/answer/2482017) | İşlem türleri belgeyle karşılaştırıldı |
| App Store Connect: Summary Sales Report (sütunlar, iade işareti) | [developer.apple.com](https://developer.apple.com/help/app-store-connect/reference/summary-sales-report) | Sütun adları ve iade gösterimi belgeyle karşılaştırıldı |

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
