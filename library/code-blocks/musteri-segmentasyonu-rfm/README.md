# Müşteri Segmentasyonu (RFM) · Kod Bloğu

> Pazarlama › Pazarlama Uzmanı · Workers / Workless

Satış geçmişinden müşterileri **yenilik (R), sıklık (F) ve tutara (M)** göre puanlar ve segmentlere ayırır. Her
segmentin müşteri ve ciro payını gösterir, segment bazlı aksiyon önerir ve kaybedilmek üzere olan müşterileri
işaretler. İstenirse her segment için CRM'e aktarılabilecek bir CSV yazar. İnternete bağlanmaz.

## Nasıl hesaplar?

- **Pencere:** Analiz tarihinden geriye `--pencere-ay` (12) ay. Pencere dışındaki ve analiz tarihinden sonraki satırlar
  alınmaz.
- **R (yenilik):** analiz tarihi − son alım tarihi (gün). Az gün iyidir.
- **F (sıklık):** farklı fatura / sipariş sayısı. Aynı fatura numaralı satırlar tek alım sayılır. Fatura numarası
  yoksa her tarih bir alımdır. İade satırları (eksi tutar) alım sayılmaz.
- **M (tutar):** net tutar = satışlar − iadeler. Net tutarı 0 veya eksi olan müşteri puanlanmaz.
- **Puan (1–5):** 1 + ⌊5 × (kendisinden kötü değere sahip müşteri oranı)⌋, en çok 5.
  - Eşit değerler aynı puanı alır. Örneğin tek alımlı tüm müşterilerin F puanı 1'dir.
  - Bu yöntem beşli dilimlemeye (quintile) yakındır, ama eşit değerli müşterileri farklı dilimlere bölmez.

### Segmentler (R ve F puanına göre)

| Segment | R | F | Önerilen aksiyon |
|---|---|---|---|
| Şampiyonlar | 5 | 4–5 | Ödüllendirin; yeni ürünleri önce bu gruba duyurun; referans isteyin |
| Sadık Müşteriler | 3–4 | 4–5 | Sadakat programı, çapraz / üst satış |
| Potansiyel Sadıklar | 4–5 | 2–3 | Sadakat programına davet, kişiselleştirilmiş öneri |
| Yeni Müşteriler | 5 | 1 | Hoş geldin iletişimi; ikinci alımı teşvik |
| Umut Vaat Edenler | 4 | 1 | Bilinirlik iletişimi; küçük ve süreli teşvik |
| İlgi Bekleyenler | 3 | 3 | Sınırlı süreli teklif, geçmiş alıma göre öneri |
| Uyumak Üzere | 3 | 1–2 | Popüler ürün önerisi, yeniden bağlantı kampanyası |
| Risk Altında | 1–2 | 3–4 | Kişisel iletişim, geri kazanma kampanyası |
| Kaybedilmemesi Gerekenler | 1–2 | 5 | Satış temsilcisi araması, özel teklif, sorun var mı kontrolü |
| Uykudakiler | 1–2 | 1–2 | Düşük maliyetli yeniden aktivasyon |

Harita 25 R × F hücresinin tamamını kapsar. M puanı raporda gösterilir, segmenti değiştirmez. Aynı segmentte M
puanına göre önceliklendirme yapabilirsiniz.

- **Kayıp sinyali:** En az 3 alımı olan ve son alımdan bu yana geçen süre, ortalama alım aralığının `--kayip-kat` (2)
  katını aşan müşteri.

| Uyarı | Önem |
|---|---|
| Kayıp sinyali olan müşteriler; "Risk Altında" ve "Kaybedilmemesi Gerekenler" segmentleri | Orta |
| 20'den az müşteri (5'li puanlama anlamlı değil); gelecek tarihli satır; okunamayan satır | Orta |
| Pencere dışı satırlar; yalnız iadesi olan veya net tutarı eksi müşteri | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek: 120 müşteri, analiz tarihi 09.10.2026
python main.py --satislar satislar.xlsx --bugun 01.10.2026 --pencere-ay 12 --segment-dosyalari
```

| Sütun | Açıklama |
|---|---|
| Müşteri No | Zorunlu (cari kod) |
| Fatura Tarihi, Tutar | Zorunlu; iade satırı eksi tutarla |
| Fatura No | Önerilir (aynı faturanın satırları tek alım) |
| Müşteri Adı, Bölge, Kanal | İsteğe bağlı |

## Çıktı

- `rfm_segmentasyonu.xlsx`:
  - `Segment Özeti`: müşteri ve ciro payı, medyan R ve F, ortalama M, önerilen aksiyon.
  - `Müşteriler`: R, F, M, iade, puanlar, RFM kodu, segment, ortalama alım aralığı, kayıp sinyali.
  - `R × F Matrisi`: segment renkleriyle.
  - `Kayıp Sinyali`: boş "Aksiyon / Sorumlu" sütunuyla.
  - `Uyarılar`.
- `--segment-dosyalari` ile `segmentler/<segment>.csv` dosyaları yazılır (UTF-8, `;` ayraçlı; Excel ve çoğu CRM
  açar).

## Dikkat

- **Puanlar görecelidir.** Aynı müşteri farklı müşteri kitlesinde farklı puan alabilir. Sektörler ve dönemler arası
  karşılaştırma yapmayın.
- **Satın alma döngüsü:** Yılda bir alım yapılan işlerde (ör. yıllık sözleşme) R ve F yorumu farklıdır. Pencereyi ve
  kayıp katsayısını işinize göre ayarlayın.
- **Kişisel veri:** Bireysel müşterilerde ad ve iletişim bilgisi kişisel veridir. Kampanya iletişimi için ticari
  elektronik ileti izni (İYS) ve KVKK aydınlatma / açık rıza yükümlülüklerini kontrol edin. Segment dosyalarını yalnız
  yetkili kişilerle paylaşın.
- **Örnek veri:** Örnek müşteriler ve satışlar kurgusaldır.

## Testler

Şunlar test edilir:
- Puan formülü: eşit değerler ve R'nin ters yönü (elle hesaplanmış).
- Segment haritasının 25 hücrenin tamamını kapsaması.
- İade, aynı fatura numaralı satırlar, pencere dışı ve gelecek tarihli satırlar.
- Örnek veride toplamların tutarlılığı, kayıp sinyali koşulu, segment CSV'leri, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
