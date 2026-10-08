# Hasar Sıklık ve Şiddet Analizi · Kod Bloğu

> Sigortacılık › Aktüerya › Aktüerya Uzmanı · Workers / Workless

Poliçe ve hasar dökümünden dönem ve segment bazında (araç tipi, bölge, ürün…) **hasar sıklığı, şiddeti, saf prim
ve hasar/prim oranı** hesaplar. Fiyatlama ve portföy analizinin temel tablosudur. İnternete bağlanmaz.

## Ne hesaplar?

| Ölçüt | Tanım |
|---|---|
| Maruziyet | Poliçe süresinin döneme düşen günü ÷ 365 (poliçe-yıl); değerleme tarihinde ve iptal tarihinde kesilir |
| Kazanılmış prim | Prim × döneme düşen gün ÷ poliçe süresi |
| Gerçekleşen hasar | Ödenen + muallak (reddedilen dosyalar hariç) |
| Sıklık | Hasar adedi ÷ maruziyet; %95 güven aralığı (Poisson yaklaşımı) |
| Şiddet | Gerçekleşen ÷ adet; **sınırlı şiddet**: büyük hasar eşiğinde sınırlanmış |
| Saf prim | Gerçekleşen ÷ maruziyet (= sıklık × şiddet) |
| Hasar/prim | Gerçekleşen ÷ kazanılmış prim |
| Güvenilirlik | √(adet ÷ 1.082); klasik tam güvenilirlik ölçütü (p = %90, k = %5); %50 altı sarı |

**Büyük hasar eşiği:** `--buyuk-hasar` ile tutar (ör. `500000`) veya yüzdelik (ör. `p95`, `p99`) olarak
verilir. Eşik üstündeki hasarlar ayrıca sayılır.

**Kırılımlar:** Kırılım sütunları `--segment` ile seçilir ve birden fazla verilebilir. Dönem yıllık veya çeyreklik
(`--donem ceyrek`) olabilir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                    # örnek kasko portföyü (kurgusal)
python main.py --policeler policeler.xlsx --hasarlar hasarlar.xlsx --segment "Araç Tipi" --segment Bölge \
               --degerleme 30.09.2026
python main.py --policeler p.xlsx --hasarlar h.xlsx --donem ceyrek --buyuk-hasar 500000
```

| Dosya | Sütunlar |
|---|---|
| Poliçeler | Poliçe No, Başlangıç, Bitiş, Prim, [İptal Tarihi], segment sütunları |
| Hasarlar | Hasar No, Poliçe No, Hasar Tarihi, Ödenen, Muallak, [Durum] — veya Gerçekleşen |

## Çıktı

`Dönemler` (sıklık ve şiddet grafiği dahil) · her segment için ayrı sayfa · `Dönem × Segment` · `Bilgi`

## Dikkat

- **Gelişmemiş dönemler:** Gerçekleşen hasar IBNR'ı içermez. Son dönemlerin hasarları henüz tam gelişmediği için
  sıklık ve şiddet düşük görünür. Fiyatlamada gelişim faktörleri ve IBNR ile düzeltin (bkz. **IBNR Rezerv Tahmini**
  paketi).
- **Enflasyon:** Yüksek enflasyon döneminde şiddet karşılaştırmaları için tutarları ortak bir tarihe taşıyın
  (trend/endeksleme).
- **Küçük segmentler:** Güvenilirliği düşük segmentlerin sonuçları dalgalıdır; komşu segmentlerle birleştirerek
  değerlendirin.

## Testler

Testler elle hesaplanmış küçük bir portföyle maruziyet, kazanılmış prim, sıklık, şiddet ve H/P değerlerini
doğrular.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
