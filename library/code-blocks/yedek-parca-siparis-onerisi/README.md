# Yedek Parça Sipariş Önerisi · Kod Bloğu

> Yedek Parça › Yedek Parça Danışmanı · Workers / Workless

Parça bazında tüketim, bekleyen iş emirleri ve tedarik süresine göre **ne kadar sipariş verilmesi gerektiğini**
önerir; acil kalemleri öne çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Günlük tüketim | Aylık ortalama tüketim ÷ `--ay-gun` (30) |
| Kullanılabilir stok | Eldeki stok + yoldaki sipariş − bekleyen iş emri ihtiyacı |
| Sipariş noktası | Günlük tüketim × (tedarik süresi + `--emniyet-gun` 7) |
| Hedef stok | Günlük tüketim × (tedarik süresi + emniyet günü + `--kapsama-gun` 30) |
| Önerilen sipariş | Kullanılabilir ≤ sipariş noktası ise hedef − kullanılabilir; paket katına yukarı yuvarlanır |
| Stok gün | Eldeki stok ÷ günlük tüketim |

Durum: **Acil** (bekleyen iş emri karşılanamıyor veya eldeki stok tedarik süresince yetmiyor), **Sipariş ver**,
**Yeterli**, **Hareketsiz** (tüketim yok, stok var).

**Örnek:** Stok 20, bekleyen iş emri 5, aylık tüketim 60 (günde 2), tedarik 10 gün, paket 10. Kullanılabilir 15;
sipariş noktası 2 × 17 = 34; hedef 2 × 47 = 94; ihtiyaç 79 → 80 adet. Eldeki 15 adet 10 günlük tüketimi (20)
karşılamadığı için "Acil".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --parcalar parcalar.xlsx
python main.py --parcalar parcalar.xlsx --emniyet-gun 10 --kapsama-gun 45 --ay-gun 26
```

Zorunlu sütunlar: Parça No, Stok, Aylık Tüketim. Diğerleri: Parça Adı, Yoldaki, Bekleyen İş Emri, Tedarik Süresi
(gün), Paket Adedi (boşsa 1), Birim Fiyat.

## Çıktı

`cikti/yedek_parca_siparis_onerisi.xlsx`:

- `Sipariş Önerisi`: acil kalemler üstte; sipariş noktası, hedef stok, önerilen adet, tutar ve açıklama.
- `Uyarılar`: acil siparişler, hareketsiz stok, mükerrer ve okunamayan satırlar.

## Dikkat

- **Ortalama tüketim sizin verinizdir.** Kod mevsimselliği (kış lastiği, akü, klima), kampanyaları, geri çağırmaları
  ve model değişimini bilmez. Ortalamayı hangi dönemden aldığınız sonucu doğrudan değiştirir.
- **Emniyet ve kapsama günleri varsayımdır.** Hızlı dönen sarf parçası ile yavaş dönen pahalı parça aynı kuralla
  yönetilmez; parça grubuna göre ayrı çalıştırın.
- **Yoldaki sipariş gelecekmiş gibi sayılır.** Gecikme riski olan siparişleri "Yoldaki" sütununa yazmayın.
- **Hareketsiz parçaya sipariş önerilmez.** Bekleyen iş emri varsa yalnız eksik kadar önerilir.
- **Üretici sipariş kuralları kapsam dışıdır.** Stok ve acil sipariş iskontoları, asgari sipariş tutarı ve iade
  koşullarını siz değerlendirirsiniz.
- **Örnek veri:** Parçalar, miktarlar ve fiyatlar kurgusaldır.

## Testler

- Sipariş noktası, hedef stok, paket katına yuvarlama, küsuratlı tüketim.
- Dört durumun sınır değerleri; emniyet, kapsama ve ay günü ayarlarının etkisi; tüketimi olmayan parçalar.
- Sıralama, toplam tutar, fiyatı olmayan kalem; okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
