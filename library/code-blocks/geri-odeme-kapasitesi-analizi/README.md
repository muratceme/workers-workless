# Geri Ödeme Kapasitesi Analizi · Kod Bloğu

> Bankacılık › Krediler (Tahsis) › Kredi Tahsis Uzmanı · Workers / Workless

Firmanın nakit yaratma gücünü mevcut ve önerilen kredilerin **borç servisiyle** karşılaştırır. Yıl bazında
**borç servisi karşılama oranını (DSCR)** ve stres senaryolarını hesaplar; taşınabilir azami kredi tutarını
bulur. İnternete bağlanmaz.

## Nasıl hesaplar?

1. **Varsayımlar (geçmiş mali tablolardan):**
   - Satış büyümesi: geçmiş bileşik büyüme ya da `--buyume`.
   - FAVÖK marjı: son yıl.
   - Vergi / FAVÖK ve yatırım / satış: geçmiş yılların toplamından.
   - Net işletme sermayesi (NİS) / satış: son yıl. NİS = ticari alacak + stok − ticari borç.
2. **Borç servisine kullanılabilir nakit (CFADS):** `FAVÖK − vergi − yatırım − NİS artışı`. Yüksek nominal
   büyümede NİS artışı nakdin önemli kısmını tüketir; bu yüzden hesaba katılır.
3. **Borç servisi:** Her kredinin aylık ödeme planından yıllık anapara + faiz hesaplanır.
   - Eşit taksit, eşit anapara ve vade sonu ödemeli krediler planına göre işlenir.
   - **Rotatif** kredilerde yalnız faiz ödenir; anaparanın yenileneceği varsayılır.
   - Önerilen kredide ödemesiz dönem boyunca faiz ödenir.
   - Dövizli krediler kurla TL'ye çevrilir.
   - Aylık faiz = yıllık nominal faiz / 12.
4. **DSCR** = CFADS / borç servisi, her projeksiyon yılı için. Projeksiyon en uzun kredinin vadesi kadar sürer.
   Net finansal borç / FAVÖK de hesaplanır.
5. **Senaryolar:**

   | Senaryo | Değişiklik |
   |---|---|
   | Baz | Varsayımlar olduğu gibi |
   | Satış −%15 | Satışlar %15 düşük |
   | FAVÖK marjı −3 puan | Marj 3 puan düşük |
   | Değişken faiz +10 puan | Değişken faizli kredilerde faiz 10 puan yüksek |
   | Kur +%30 | Döviz kredilerinin TL servisi %30 yüksek |
   | Birleşik stres | Satış −%10, marj −2 puan, faiz +5 puan, kur +%20 |

   **Sonuç etiketi:**
   - En düşük DSCR ≥ eşik (`--min-dscr`, varsayılan 1,25): **Karşılıyor**.
   - 1,0 ≤ DSCR < eşik: **Sınırda**.
   - DSCR < 1,0: **Karşılamıyor**.
6. **Azami kredi ve kırılma noktası:**
   - Aynı vade, faiz ve ödeme yapısıyla baz senaryoda eşiği sağlayan en yüksek kredi tutarı.
   - En düşük DSCR'yi 1,0'a indiren satış düşüşü.

## Kullanım

```bash
pip install -r requirements.txt
python main.py                                                          # örnek firma, 20 milyon TL kredi önerisi
python main.py --mali mali_tablolar.xlsx --borclar borclar.xlsx --oneri kredi_onerisi.json --min-dscr 1,2
python main.py --mali mali.xlsx --borclar borclar.xlsx --oneri oneri.json --buyume 25 --kur EUR=48,05 USD=41,20
```

| Dosya | İçerik |
|---|---|
| Mali tablolar | İlk sütun kalem, sonraki sütunlar yıllar. **Net Satışlar, FAVÖK** zorunlu; Ödenen Vergi, Yatırım Harcamaları, Ticari Alacaklar, Stoklar, Ticari Borçlar, Hazır Değerler |
| Mevcut krediler | Kredi, Banka, Bakiye, Para Birimi, Yıllık Faiz (%), Kalan Vade (Ay), Ödeme Tipi (Eşit taksit / Eşit anapara / Vade sonu / Rotatif), Faiz Tipi (Sabit / Değişken) |
| Kredi önerisi (JSON) | `tutar`, `para_birimi`, `yillik_faiz`, `vade_ay`, `odemesiz_ay`, `odeme_tipi`, `faiz_tipi`, `kur` |

Örnek sonuç: Baz senaryoda en düşük DSCR 1,14 (Sınırda). Kur şoku ve birleşik streste karşılamıyor. Aynı
yapıyla taşınabilir azami kredi yaklaşık 16,4 milyon TL.

## Çıktı

`geri_odeme_kapasitesi.xlsx`:
- `Özet`.
- `Projeksiyon (Baz)`: CFADS, borç servisi, DSCR ve net borç/FAVÖK, grafikli.
- `Senaryolar`: yıl bazında DSCR, grafikli.
- `Borç Servisi`: kredi bazında yıllık anapara ve faiz.
- `Ödeme Planı`: önerilen kredinin aylık planı.
- `Varsayımlar`.

## Dikkat

- **Basitleştirilmiş projeksiyon:** Aşağıdakiler dikkate alınmaz:
  - Faiz giderinin vergi etkisi.
  - Kredinin finanse ettiği yatırımın getirisi.
  - Kâr payı dağıtımı, ortak hareketleri ve yeni borçlanmalar.
  - Kredi maliyetine eklenen BSMV, komisyon ve masraflar. Gerekirse faiz oranına ekleyin.
- **Nominal büyüme ve enflasyon:** Büyüme nominaldir; yüksek enflasyon döneminde geçmiş büyüme geleceği
  olduğundan iyi gösterebilir. Muhafazakâr bir `--buyume` ile de deneyin.
- **Kredi kararı:** DSCR eşiği ve senaryolar örnektir. Kredi kararı, iç derecelendirme ve fiyatlama bankanın
  yetkili organlarına ve politikalarına aittir.
- **Örnek veri:** Örnek firma ve krediler kurgusaldır.

## Testler

Testler şunları elle hesaplanan değerlerle karşılaştırır:
- Annüite ödemesi: 100.000 TL, aylık %1, 12 ay → 8.884,88 TL.
- Ödemesiz dönem.
- İlk yılın CFADS ve borç servisi.
- Azami kredinin eşiği sağladığı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçlar kontrol edilmeden işlem yapılmamalıdır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
