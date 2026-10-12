# Proje Maliyet Kontrol Raporu · Kod Bloğu

> Teknik Ofis › Teknik Ofis Şefi · Workers / Workless

Bütçe, gerçekleşen maliyet ve hakediş gelirlerinden **iş kalemi bazında maliyet sapması ve tamamlanma tahmini** üretir.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Kazanılan değer | Bütçe × fiziksel tamamlanma % |
| Maliyet sapması | Kazanılan değer − gerçekleşen maliyet (eksi: yapılan iş için bütçeden fazla harcanmış) |
| Maliyet performansı | Kazanılan değer ÷ gerçekleşen maliyet |
| Tamamlanma tahmini | `performans` (varsayılan): gerçekleşen × bütçe ÷ kazanılan değer. `butce`: gerçekleşen + (bütçe − kazanılan değer) |
| Kalan maliyet | Tamamlanma tahmini − gerçekleşen maliyet |
| Tamamlanma sapması | Bütçe − tamamlanma tahmini (eksi: bütçe aşımı bekleniyor) |
| Bugüne kadarki brüt kâr | Hakediş geliri − gerçekleşen maliyet |

Tamamlanmış kalemde tahmin gerçekleşen maliyettir; ilerlemesi veya maliyeti olmayan kalemde bütçe yöntemi kullanılır.

| Durum | Koşul |
|---|---|
| Başlamadı | Maliyet ve ilerleme yok |
| Bütçe aşıldı | Gerçekleşen maliyet > bütçe |
| Aşım bekleniyor | Tamamlanma sapması < −`--tolerans` (varsayılan %5) |
| Bütçe altında | Tamamlanma sapması > +tolerans |
| Bütçede | Diğerleri |

**Örnek:** bütçe 10.000.000, gerçekleşen 6.600.000, tamamlanma %60 → kazanılan değer 6.000.000; maliyet sapması
−600.000; tahmin 6.600.000 × 10.000.000 ÷ 6.000.000 = 11.000.000; sapma −1.000.000 (−%10) → Aşım bekleniyor.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (sözleşme bedeli 28.000.000)
python main.py --kalemler is_kalemleri.xlsx --yontem butce --tolerans 3 --sozlesme-bedeli 28000000
```

Sütunlar: İş Kalemi, Bütçe, Gerçekleşen Maliyet, Tamamlanma % (zorunlu); Hakediş Geliri.

## Çıktı

`cikti/maliyet_kontrol_raporu.xlsx` → `Maliyet Kontrol` (sorunlu kalemler başta; tüm göstergeler, toplam satırı, sözleşme
bedeli verildiyse tahmini proje kârı; boş "Açıklama / Aksiyon" sütunu) ve `Uyarılar`.

## Dikkat

- **Tahmin, tamamlanma yüzdesine dayanır.** Yüzde iyimser girilirse sapma küçük görünür. Yüzdeyi ölçülen imalat
  miktarından (metraj, hakediş) türetin; harcanan paradan veya geçen süreden değil.
- **Taahhütler dahil değildir.** Sipariş verilmiş ama faturası gelmemiş malzeme, taşeron hakedişi ve kesin hesap
  farkları "gerçekleşen maliyet"e eklenmediyse durum olduğundan iyi görünür.
- **İki yöntem iki farklı varsayımdır.** `performans` bugünkü verimin sonuna kadar süreceğini, `butce` kalan işin
  bütçesinde biteceğini varsayar; gerçek çoğu zaman ikisinin arasındadır.
- **Bütçe ve maliyet aynı esasta olmalı.** KDV, genel gider ve fiyat farkı iki tarafta da aynı şekilde yer almalıdır.
- **Süre boyutu yoktur.** İş programı gecikmesi için "Kazanılmış Değer Analizi" paketine bakın.
- **Örnek veri:** İş kalemleri ve tutarlar kurgusaldır.

## Testler

- Kazanılan değer, sapmalar, iki tahmin yöntemi, tamamlanmış ve başlamamış kalem, durum sınırları (tam %5).
- Toplamlar, proje tamamlanma yüzdesi, sözleşme bedeliyle tahmini kâr, uyarılar.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
