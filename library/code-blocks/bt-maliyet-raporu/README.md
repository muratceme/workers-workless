# BT Maliyet Raporu · Kod Bloğu

> Bilgi Teknolojileri › BT Müdürü · Workers / Workless

Lisans, bulut, donanım, hizmet ve iletişim harcamalarını birim, kategori ve tedarikçi bazında raporlar. Bütçe
sapmasını, atıl lisansları, mükerrer abonelikleri ve yaklaşan yenilemeleri işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

- **Kategori:** Dosyada yoksa açıklamadan anahtar kelimeyle bulunur. Sıra: bulut → iletişim → donanım → lisans →
  hizmet. Örneğin "Bulut sunucu" Bulut, "Sunucu disk yükseltme" Donanım, "ERP lisans bakım" Lisans sayılır.
- **TL karşılığı:** tutar × kur. Kur yoksa dosyadaki aynı dövizin en yakın tarihli kuru kullanılır ve uyarı verilir.
- **Tahakkuk:**
  - "Yıllık" ödemeler başlangıç ayından itibaren 12 aya yayılır. Rapor döneminden sonraya düşen kısım "peşin ödenmiş,
    sonraki aylara ait" olarak ayrı gösterilir.
  - "Aylık" ve "Tek seferlik" ödemeler ödendiği aya yazılır.
  - Tahakkuk + peşin kısım = ödenen toplam.
- **Bütçe:** Kategori yıllık bütçesi, rapor dönemi için oransal bütçe ve yıl sonu tahmini karşılaştırılır.
  Tahmin = aylık kalemlerin ortalaması × 12 + yıllık ödemeler + tek seferlik harcamalar.
- **Kişi başı maliyet:** Birim çalışan sayıları verilirse hesaplanır. "Genel" birim (ör. şirket geneli lisanslar)
  dağıtılmaz, ayrı gösterilir.
- **Lisans kullanımı:** Atıl lisans = toplam − kullanılan; atıl lisans maliyeti = atıl × yıllık birim fiyat.

| Kontrol | Önem |
|---|---|
| Kuru olmayan ve tahmin de edilemeyen döviz harcaması; lisans aşımı (kullanıcı > lisans); yıl sonu tahmini bütçeyi %10'dan fazla aşıyor | Yüksek |
| Yıl sonu tahmini bütçeyi aşıyor | Orta |
| Aylık kalemde önceki 3 ay ortalamasından %30 (`--artis`) fazla artış (ör. bulut) | Orta |
| Farklı birimlerin aynı aboneliği ayrı ayrı alması (olası mükerrer) | Orta |
| Lisans kullanımı %80'in altında; kur tahmini | Orta |
| `--yenileme-gun` (120) içinde biten sözleşme / lisans | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 49 harcama, Ocak–Eylül 2026, rapor tarihi 09.10.2026
python main.py --harcamalar h.xlsx --lisanslar l.xlsx --butce b.csv --birimler birimler.csv --yil 2026 --bugun 09.10.2026
```

| Dosya | Sütunlar |
|---|---|
| Harcamalar | Tarih, Tedarikçi, Kalem, Kategori (isteğe bağlı), Birim, Tutar, Döviz, Kur, Dönem Tipi, Sözleşme Bitiş |
| Lisanslar (isteğe bağlı) | Ürün, Toplam Lisans, Kullanılan Lisans, Yıllık Birim Fiyat, Döviz, Yenileme Tarihi |
| Bütçe (isteğe bağlı) | Kategori, Yıllık Bütçe |
| Birimler (isteğe bağlı) | Birim, Çalışan |

## Çıktı

`bt_maliyet_raporu.xlsx`:
- `Özet`: tahakkuk, ödenen, peşin kısım, aylık ortalama, döviz payı; birim bazında kişi başı maliyet.
- `Kategori × Ay`: grafikli.
- `Birim × Kategori`.
- `Tedarikçiler`: pay ve kümülatif pay (Pareto).
- `Bütçe`: oransal bütçe, yıl sonu tahmini, fark; boş "Açıklama" sütunu.
- `Lisans Kullanımı`: boş "Karar" sütunuyla.
- `Yenileme Takvimi`: boş "Aksiyon" sütunuyla.
- `Harcamalar`: TL karşılıklı tüm satırlar.
- `Uyarılar`.

## Dikkat

- **Kategori kuralları:** Anahtar kelimeyle kategori bulmak yanılabilir. Önemli kalemlerde "Kategori" sütununu
  doldurun.
- **Kur:** Muhasebe kaydındaki fatura kurunu kullanın. Tahmini kur yalnız ön rapor içindir.
- **Tahmin:** Yıl sonu tahmini mevcut harcama hızına dayanır. Planlanmış yeni alımlar dahil değildir.
- **Örnek veri:** Örnek harcamalar ve lisanslar kurgusaldır.

## Testler

Şunlar test edilir:
- Kategori kuralları; yıllık ödemelerin tahakkuku (yıl içi ve gelecek yıla kalan).
- Tahakkuk + peşin = ödenen korunumu.
- Artış, mükerrer abonelik, atıl lisans (tam %80 sınırı), kur tahmini, bütçe ve yenileme uyarıları.
- Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
