# Sipariş Termin Planı (T&A) · Kod Bloğu

> Tekstil ve Konfeksiyon › Müşteri Temsilciliği (Merchandising) › Production Merchandiser · Workers / Workless

Sipariş sevk (ex-factory) tarihinden **geriye doğru** Time & Action (T&A) planı çıkarır. Planda lab dip, kumaş,
aksesuar, PP numune, kesim, dikim, ütü-paket ve final kontrol aşamaları yer alır. Gerçekleşen tarihlerle
karşılaştırarak geciken ve gecikecek aşamaları ve **tahmini sevk gecikmesini** gösterir. İnternete bağlanmaz.

## Nasıl çalışır?

1. **Şablon:** `ta_sablonu.csv` aşamaları, sürelerini ve hangi aşamadan önce bitmeleri gerektiğini tanımlar.
   - **Süre türü:** Sabit iş günü (kumaş üretimi 25 gün) veya adete bağlı olabilir. Adete bağlı süre, adet ÷ günlük
     kapasite olarak hesaplanır (dikim 900 adet/gün).
   - **Sipariş bazında kapasite:** Sipariş dosyasında `Kapasite Dikim` gibi bir sütunla aşama kapasitesi
     değiştirilebilir.
2. **Geriye planlama:** Her aşamanın **en geç bitişi**, sonraki aşamanın en geç başlangıcından önceki iş günüdür.
   Sayılmayan günler:
   - Pazar (`--calisma-gunu 5` ile Cumartesi de)
   - sabit genel tatiller ve dini bayramlar
3. **Plan payı:** İlk aşamaların en geç başlangıcı ile PO tarihi arasındaki iş günü farkıdır. Negatifse sipariş bu
   şablonla yetişmez; ya süreler kısaltılmalı ya termin konuşulmalıdır.
4. **İleri tahmin:** Gerçekleşen tarihler verilirse tahmini bitişler şöyle hesaplanır:
   - Biten aşama gerçekleşen tarihte biter.
   - Açık aşama öncülleri bittikten sonra başlar ve süresi kadar sürer; bugünden önce bitmiş sayılamaz.
   - Sevkin tahmini bitişi ile son tarihi arasındaki fark **tahmini sevk gecikmesidir**.

**Durumlar:** Tamamlandı · Geç tamamlandı · **GECİKTİ** (son tarihi geçti, bitmedi) · **Gecikecek** (öncül
gecikmesi nedeniyle) · Yaklaşıyor (`--uyari-gun`, varsayılan 3 iş günü) · Planlandı

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                    # örnek: 3 sipariş, durum tarihi 08.10.2026
python main.py --siparisler siparisler.xlsx --sablon ta_sablonu.csv --gerceklesen gerceklesen.xlsx
python main.py --siparisler siparisler.xlsx --bugun 15.10.2026 --calisma-gunu 5
```

| Dosya | Sütunlar |
|---|---|
| Siparişler | Sipariş No, Müşteri, Model, Adet, PO Tarihi, Sevk Tarihi, [Kapasite Dikim, Kapasite Kesim …] |
| Şablon | Aşama, Süre (gün), Sonraki Aşama, Süre Türü (sabit/adet), Kapasite (adet/gün) |
| Gerçekleşen | Sipariş No, Aşama, Gerçekleşen Tarih |

**Şablon kuralları**
- Sonraki aşaması boş olan tek bir aşama bulunmalıdır; bu aşama sevktir.
- Bir aşamadan önce birden fazla aşama bitebilir. Örneğin kesimden önce kumaş, aksesuar ve PP numunesinin
  tamamlanması gerekir.

## Çıktı

`Sipariş Özeti` (riske göre sıralı) · `T&A Matrisi` (sipariş × aşama son tarihleri, renkli durum) · `Aşama Detayı` ·
`Bilgi`

## Dikkat

- **Şablon süreleri:** Örnektir. Kendi kumaşçı, boyahane, aksesuar ve atölye sürelerinizle güncelleyin.
- **İhtiyatlı plan:** Kesim, dikim ve ütü-paket uygulamada kısmen paralel yürür. Bu araç aşamaları sırayla
  (tamponlu) planlar, bu yüzden ihtiyatlı bir plan verir.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
