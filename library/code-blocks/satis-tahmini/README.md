# Satış Tahmini · Kod Bloğu

> Satış › Satış Müdürü · Workers / Workless

Geçmiş aylık satışlardan **ürün grubu ve bölge bazında** aylık satış tahmini üretir. Her seri için birkaç yöntemi
geçmiş veride dener ve en az hata yapanı seçer. Mevsimselliği, aykırı ayları ve tahmin güvenini raporlar. İnternete
bağlanmaz.

## Nasıl hesaplar?

1. **Seriler:** Ürün Grubu × Bölge bazında aylık toplam. Seri içinde eksik ay 0 sayılır.
2. **Yöntemler:**

| Yöntem | Açıklama | Gereken veri |
|---|---|---|
| Naif | Son ayın değeri | 1 ay |
| Mevsimsel naif | Geçen yılın aynı ayı | 12 ay |
| Hareketli ortalama | Son 3 ayın ortalaması | 3 ay |
| Holt | Düzey + doğrusal trend | 3 ay |
| Holt-Winters toplamsal | Düzey + trend + 12 aylık mevsimsellik (sabit genlik) | 24 ay |
| Holt-Winters çarpımsal | Düzey + trend + mevsimsellik (satışla orantılı genlik) | 24 ay, sıfır olmayan değerler |

   Düzeltme katsayıları (α, β, γ ∈ {0,1; 0,3; 0,5; 0,7; 0,9}), tek adımlı tahmin hatalarının kareler toplamını en
   küçükleyecek şekilde seçilir.
3. **Geriye dönük test:** Son `--test` (6) ay dışarıda bırakılır ve her yöntem önceki verilerle bu ayları tahmin eder.
   - En düşük **WAPE**'li yöntem seçilir: WAPE = Σ|gerçek − tahmin| / Σ gerçek.
   - Seçilen yöntem tüm veriyle yeniden kurulur ve `--ufuk` (12) ay tahmin edilir. Eksi tahmin 0 yapılır.
4. **Yaklaşık aralık:** tahmin ± 1,28 × test RMSE × √h, yaklaşık %80'lik aralıktır; h ileri adım sayısıdır.
5. **Mevsimsellik endeksi:** Ürün grubunun ay ortalaması / tüm ayların ortalaması × 100.

| Uyarı | Önem |
|---|---|
| Testte WAPE > %30 (düşük güven) | Orta |
| Aykırı ay: geçen yılın aynı ayına oran, tüm ayların medyan oranından 4 MAD'den fazla sapıyor | Orta |
| Seri diğerlerinden önce bitiyor (son aylar eksik); test için yetersiz kısa seri | Orta |
| 24 aydan kısa seri (mevsimsel yöntem denenmez); ayların %30'undan fazlası sıfır (seyrek satış) | Bilgi |

Bir ay aykırı işaretlenince, ertesi yılın aynı ayı ondan etkilendiği için tekrar işaretlenmez.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                        # örnek: 3 ürün grubu × 4 bölge, Ekim 2023 – Eylül 2026
python main.py --satislar satislar.xlsx --ufuk 12 --test 6
python main.py --satislar satislar.xlsx --deger Tutar
```

| Sütun | Açıklama |
|---|---|
| Ay | 2026-09, 09.2026, Eylül 2026 veya tarih |
| Ürün Grubu, Bölge | Boşsa "Tümü" |
| Miktar / Tutar | `--deger` ile seçilen sütun |

## Çıktı

`satis_tahmini.xlsx`:
- `Tahmin`: seri × gelecek aylar, toplam, yaklaşık aralık, seçilen yöntem ve test WAPE'si.
- `Yöntem Karşılaştırma`: her yöntemin test WAPE'si; seçilen yeşil, parametreler.
- `Toplam`: gerçekleşen ve tahmin edilen toplam (grafik), son 12 ay ve gelecek 12 ay karşılaştırması.
- `Mevsimsellik`: ürün grubu × ay endeksi.
- `Geçmiş + Tahmin`: filtrelenebilir uzun tablo (alt / üst sınırlarla).
- `Uyarılar`.

## Dikkat

- **Tahmin geçmişin tekrarına dayanır.** Fiyat değişikliği, kampanya, yeni ürün, rakip hamlesi veya ekonomik şok gibi
  olaylar modelde yoktur. Sonuçları satış ekibinin bilgisiyle düzeltin.
- **Tutar tahmini enflasyondan etkilenir.** Fiyat artışlarının yüksek olduğu dönemlerde miktar tahmin edip fiyat
  varsayımıyla tutara çevirmek daha sağlıklıdır.
- **Aykırı aylar** (tek seferlik büyük sipariş, stoksuzluk) tahmini bozar. Uyarıda listelenen ayları düzeltip yeniden
  çalıştırın.
- **Örnek veri:** Örnek satışlar kurgusaldır.

## Testler

Şunlar test edilir:
- Naif, hareketli ortalama ve mevsimsel naif yöntemleri; Holt'un doğrusal seriyi tam tahmin etmesi.
- Holt-Winters'ın bilinen mevsimsel seriyi naif yöntemden çok daha iyi tahmin etmesi; 24 ay ve sıfır kontrolü.
- WAPE ve ay okuma.
- Örnek veride en düşük WAPE seçimi, toplam tutarlılığı, yaz zirvesi.
- Aykırı ay ve yankı kuralı, kısa seri, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
