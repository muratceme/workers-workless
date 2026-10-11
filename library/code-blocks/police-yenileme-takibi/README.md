# Poliçe Yenileme Takibi · Kod Bloğu

> Teknik ve Operasyon › Sigorta Uzmanı · Workers / Workless

Vadesi yaklaşan ve süresi yeni dolmuş poliçeleri listeler; **müşteri başına tek bir yenileme hatırlatma mesajı
taslağı** hazırlar. İnternete bağlanmaz ve mesaj göndermez.

## Nasıl hesaplar?

```
Kalan gün = bitiş tarihi − rapor tarihi
Kapsam    = −(--gecmis-gun) ≤ kalan gün ≤ --gun          (varsayılan: 15 gün önce bitmiş – 30 gün içinde bitecek)
```

| Grup | Kalan gün |
|---|---|
| Süresi geçmiş | 0'dan küçük (yenilenmemiş, teminat kesilmiş olabilir) |
| Bu hafta | 0 – 7 |
| Yaklaşan | 8 ve üzeri |

- `Yenilendi` sütunu `Evet` olan poliçeler alınmaz.
- **Mesaj:** aynı müşterinin (ad + telefon) kapsamdaki tüm poliçeleri tek mesajda toplanır. Her poliçe için branş,
  poliçe numarası, konu (ör. plaka), sigorta şirketi, bitiş tarihi ve kalan gün yazılır.
- **Şablon:** `--sablon` ile verdiğiniz metin dosyasında `{musteri}`, `{policeler}` ve `{imza}` alanları kullanılır.
  `--imza` verilmezse müşteri temsilcisinin adı yazılır.
- **İletişim izni** `Hayır` olan müşterinin mesajı yine hazırlanır ama işaretlenir.

Varsayılan mesaj örneği:

```
Sayın Ayşe Örnek,

- Kasko poliçeniz (Y-01, 34 ABC 001), Kurgu Sigorta: 05.10.2026 tarihinde sona eriyor (4 gün kaldı).
- Trafik poliçeniz (Y-02, 34 ABC 001), Kurgu Sigorta: 05.10.2026 tarihinde sona eriyor (4 gün kaldı).

Teminatınızın kesintiye uğramaması için yenileme işlemlerinizi birlikte tamamlayabiliriz. ...
```

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 01.10.2026)
python main.py --policeler policeler.xlsx --gun 45 --gecmis-gun 7 --imza "Acente Adı · 0212 000 00 00"
python main.py --sablon mesaj_sablonu.txt
```

Sütunlar: Poliçe No, Müşteri, Bitiş Tarihi (zorunlu); Telefon, E-posta, Branş, Sigorta Şirketi, Konu, Prim, Müşteri
Temsilcisi, İletişim İzni, Yenilendi.

## Çıktı

`cikti/yenileme_takibi.xlsx`:

- `Yenileme Listesi`: kalan güne göre sıralı poliçeler; boş "Arandı / Sonuç" sütunu.
- `Mesajlar`: müşteri başına mesaj taslağı, iletişim bilgisi ve notlar; boş "Gönderildi" sütunu.
- `Temsilci Özeti`: temsilci bazında poliçe sayısı ve prim.
- `Uyarılar`: okunamayan tarihler, mükerrer poliçeler.

`cikti/yenileme_mesajlari.txt`: mesaj taslakları, kopyalanmaya hazır.

## Dikkat

- **Mesaj göndermez.** Taslakları okuyup kendi kanalınızdan siz gönderirsiniz. Göndermeden önce poliçe bilgisini kontrol edin.
- **İletişim izni:** Ticari elektronik ileti kuralları (İYS) ve KVKK kapsamında, hatırlatmanın hangi kanaldan ve
  hangi içerikle gönderilebileceği sizin sorumluluğunuzdadır. Kod izin sütununu yalnız işaretler; izin sorgulamaz.
  Mesaja kampanya veya yeni ürün teklifi eklerseniz ticari ileti niteliği kazanabilir.
- **Fiyat içermez.** Yenileme primi teklifi hesaplanmaz; "Prim" sütunu yalnız takip amaçlı önceki primdir.
- **Zorunlu sigortalar:** Süresi geçen zorunlu sigortalarda (ör. trafik) gecikmenin sonuçları olabilir; "Süresi
  geçmiş" grubunu öncelikle arayın.
- **Kişisel veri:** Liste telefon ve e-posta içerir; çıktıları sınırlı erişimle saklayın.
- **Örnek veri:** Müşteriler, telefonlar, e-postalar ve poliçeler kurgusaldır.

## Testler

- Kapsam sınırları (−15, 0, 7, 8, 30 gün), yenilenmiş poliçeler.
- Mesaj metni: çok poliçeli müşteri, süresi geçmiş, bugün biten, imza.
- İletişim izni, özel şablon, aynı adlı farklı müşteriler; temsilci özeti.
- Okuma hataları, geçersiz şablon ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
