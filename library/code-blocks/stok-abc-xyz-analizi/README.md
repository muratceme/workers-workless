# Stok ABC-XYZ Analizi · Kod Bloğu

> Üretim › Üretim Planlama · Workers / Workless

Stokları iki eksende sınıflandırır ve her sınıf için **stok politikası** önerir. İnternete bağlanmaz.

| Eksen | Ölçüt | Varsayılan sınırlar |
|---|---|---|
| **ABC** (değer) | Tüketim değeri = miktar × birim maliyet, Pareto sıralaması | A ≤ %80, B ≤ %95, C kalan (kümülatif pay) |
| **XYZ** (talep değişkenliği) | Değişim katsayısı CV = standart sapma ÷ ortalama | X ≤ 0,5 · Y ≤ 1,0 · Z > 1,0 |

- **Sınırı aşan kalem:** Kümülatif payı sınırı aşan ilk kalem üst sınıfta kalır; örneğin %78'den %83'e çıkaran
  kalem A olur.
- **Yeni kalemler:** İlk tüketimden önceki dönemler sıfır talep sayılmaz. Bu sayede yeni kalem yapay olarak
  "düzensiz" görünmez.
- **Hareketsiz:** Hiç tüketimi olmayan kalemler ayrı listelenir.
- **Yetersiz geçmiş:** `--en-az-donem` (varsayılan 6) dönemden az geçmişi olan kalemler XYZ yerine "?"
  işaretiyle gösterilir.

Dokuz sınıfın (AX … CZ) politika önerileri rapordadır. Örnek: AX → sürekli yenileme ve düşük emniyet stoğu;
AZ → siparişe göre alım; CZ → stok tutma gerekliliğini sorgulama.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                    # örnek 12 aylık tüketimle
python main.py --girdi tuketim.xlsx --maliyet maliyetler.xlsx
python main.py --girdi tuketim.xlsx --abc 70 90 --xyz 0.25 0.5     # kendi sınırlarınız
```

ERP'den alınan **stok çıkış hareketleri** (Tarih, Stok Kodu, Miktar) doğrudan kullanılabilir; aylara toplanır.

## Çıktı

`Matris` (3×3 tablo, politika listesi, Pareto grafiği) · `Kalemler` · `Dönem Tüketimi` · `Bilgi`

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
