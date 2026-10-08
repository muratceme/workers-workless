# Satış Hedef-Gerçekleşme Raporu · Kod Bloğu

> Satış › Satış Destek Uzmanı · Workers / Workless

Satış hareketlerini aylık hedeflerle karşılaştırır. Temsilci, bölge ve ürün grubu bazında hedef-gerçekleşme,
**geçen yıl aynı dönem** ve **yıl başından bugüne** raporu üretir. Ay içinde çalıştırılırsa ay sonu tahmini ve
hedef için gereken günlük satışı da verir. İnternete bağlanmaz.

## Ne hesaplar?

| Ölçüt | Tanım |
|---|---|
| Gerçekleşme % | ay gerçekleşen ÷ ay hedefi |
| GY'ye göre % | geçen yıl **aynı gün sayısına kadar** olan satışa göre değişim (adil karşılaştırma) |
| YTD % | yıl başından bugüne gerçekleşen ÷ aynı dönemin hedef toplamı |
| Bugüne beklenen | ay hedefi × geçen iş günü ÷ ayın iş günü |
| Ay sonu tahmini | gerçekleşen ÷ geçen iş günü × ayın iş günü (doğrusal run-rate) |
| Gereken günlük | (ay hedefi − gerçekleşen) ÷ kalan iş günü |
| Prim çarpanı | `--prim 90=0.5,100=1,110=1.5` ile kapanmış ay için kademe |

**Kırılımlar**
- **Hangi sütunlar kullanılır:** Satış dosyasında hangi kırılım sütunları varsa onlar kullanılır: Temsilci,
  Bölge, Ürün Grubu ve Temsilci × Ürün Grubu.
- **Hedef gösterimi:** Hedefler yalnız hedef dosyasında verildiği boyutlarda gösterilir. Diğer kırılımlarda
  yalnız gerçekleşen ve geçen yıl karşılaştırması yer alır.
- **İadeler:** `Tür` sütununda "İade" yazan satırlar düşülür.

**Renkler:** yeşil ≥ %100 · sarı %90–99 · kırmızı < %90

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                       # örnek: 4 temsilci, Ekim 2026 ay içi (15.10)
python main.py --satislar satislar.xlsx --hedefler hedefler.xlsx --ay 2026-09 --prim 90=0.5,100=1,110=1.5
python main.py --satislar satislar.xlsx --hedefler hedefler.xlsx --ay 2026-10 --bugun 15.10.2026 --calisma-gunu 6
```

| Dosya | Sütunlar |
|---|---|
| Satışlar | Tarih, Net Tutar (KDV hariç), [Temsilci, Bölge, Ürün Grubu, Tür] — ERP satış faturası dökümü |
| Hedefler | Ay, [Temsilci, Bölge, Ürün Grubu], Hedef — veya Temsilci, 2026-01, 2026-02 … (geniş biçim) |

## Çıktı

`Genel` (aylık trend grafiği dahil) · `Temsilci` · `Bölge` · `Temsilci × Ürün Grubu` · `Bilgi`

## Dikkat

- **İş günü hesabı:** Resmî tatiller hariç tutulmaz. Bayram içeren aylarda run-rate tahmini iyimser olabilir.
- **Mevsimsellik:** Ay sonu tahmini doğrusaldır; ay sonu yoğunlaşan satışlarda gerçekleşmenin altında kalır.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
