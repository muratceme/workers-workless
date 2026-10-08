# Fason Atölye İş Takibi · Kod Bloğu

> Tekstil ve Konfeksiyon › Üretim ve Fason Takip › Fason Takip Sorumlusu · Workers / Workless

Fason atölyelere çıkan işleri takip eder: dikim, baskı, nakış, yıkama, ütü-paket. Her işin adedini, çıkış ve dönüş
tarihini ve kalite durumunu izler. Geciken işleri, kalite sorunlarını ve eksik dönüşleri listeler. Atölye
performansını ve hakedişi çıkarır. İnternete bağlanmaz.

## Ne kontrol eder?

| Kontrol | Önem |
|---|---|
| Termini geçmiş, dönmemiş adedi olan iş | Yüksek |
| Hatalı adet oranı eşiğin 2 katı veya üstü (varsayılan eşik %3, `--hata-esik`) | Yüksek |
| "Kapandı" işaretli işte eksik dönen adet fire toleransını aşıyor (varsayılan %1, `--fire-tolerans`) | Yüksek |
| Hatalı oranı eşiğin üstünde | Orta |
| Termine 2 gün veya daha az kaldı, açık adet var | Orta |
| Tempo düşük: sürenin en az yarısı geçti, dönen adet oranı geçen süre oranının 30 puan gerisinde | Orta |
| "Tamir" sonucuyla ayrılan adet; çıkandan fazla dönüş; tanımsız iş; tarih / adet hatası; fiyat yok | Orta |
| Tolerans içi fire; termin yok | Bilgi |

**Kurallar:**
- **İş durumu:**
  - **Tamamlandı:** Dönen adet çıkan adede ulaştıysa veya iş "Kapandı" işaretliyse. Son dönüş termini geçtiyse
    "Geç tamamlandı" olur.
  - **Açık adet:** çıkan − dönen.
- **Hakediş:**
  - **Hesap:** Her dönüş için sağlam adet (dönen − hatalı) × birim fiyat.
  - **Dönem:** `--donem 2026-09` yalnız o ayın dönüşlerini alır.
  - **Kapsam dışı:** KDV, stopaj ve kayıp adet mahsubu hesaba katılmaz.
- **Tamir:** Tamire ayrılan adetler atölyeye geri gider. Tamirden dönenleri yeni bir dönüş satırı olarak girin.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 5 atölye, 10 iş, durum tarihi 08.10.2026
python main.py --isler fason_isler.xlsx --donusler donusler.xlsx
python main.py --isler i.xlsx --donusler d.xlsx --bugun 15.10.2026 --hata-esik 2,5 --fire-tolerans 0,5 --donem 2026-09
```

| Dosya | Sütunlar |
|---|---|
| Fason işler | İş No, Atölye, Sipariş No, Model, İşlem, Çıkış Tarihi, Çıkan Adet, Termin, Birim Fiyat (TL), Kapandı (Evet / boş) |
| Dönüşler | İş No, Dönüş Tarihi, Dönen Adet, Kalite Sonucu (Kabul / Tamir / Red), Hatalı Adet, Hata Açıklaması |

Kısmi dönüşlerin her biri ayrı satırdır.

## Çıktı

`fason_takip.xlsx`:
- `Açık İşler`: termine göre sıralı; açık adet, hata oranı, gecikme ve boş "Atölyeyle Görüşme / Not" sütunu.
- `Tüm İşler` ve `Dönüşler` (filtrelenebilir).
- `Atölye Performansı`: atölye bazında şunlar:
  - açık iş ve atölyedeki adet
  - zamanında biten oranı
  - ortalama gecikme
  - hata oranı
  - kapanan işlerdeki kayıp adet
- `Hakediş`: dönüş bazında sağlam adet ve tutar; atölye toplamları; boş "Onay" sütunu.
- `Uyarılar`.

## Dikkat

- **Eşikler örnektir:** Hata eşiği ve fire toleransı örnektir. Atölyeyle yaptığınız sözleşmedeki kabul kriterlerini
  ve fire oranını girin.
- **Kalite kontrolü:** Dönüş kontrolü örneklemeyle yapılıyorsa hatalı adet tahmini olabilir. AQL örneklem planı için
  ayrı paketi kullanabilirsiniz.
- **Hakediş:** Ödeme öncesinde sözleşme, irsaliye ve fatura ile karşılaştırılmalıdır.
- **Örnek veri:** Örnek atölyeler ve işler kurgusaldır.

## Testler

Şunlar test edilir:
- İş durumları ve gecikme.
- Hata oranı, tamir, eksik ve fazla dönüş, tempo.
- Atölye performansı; dönem filtreli hakediş; Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
