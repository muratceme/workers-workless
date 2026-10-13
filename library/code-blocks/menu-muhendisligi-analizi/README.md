# Menü Mühendisliği Analizi · Kod Bloğu

> Yiyecek ve İçecek › Yiyecek İçecek Müdürü · Workers / Workless

Satış adetleri ve katkı paylarından menü kalemlerini **yıldız, at, bulmaca ve köpek** olarak sınıflandırır
(Kasavana-Smith menü mühendisliği yöntemi) ve her sınıf için aksiyon önerir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Katkı payı | Satış fiyatı (KDV hariç) − porsiyon maliyeti |
| Menü karışımı % | Kalemin satış adedi ÷ kategorinin toplam adedi × 100 |
| Popülerlik eşiği % | (100 ÷ kategorideki kalem sayısı) × `--esik` ÷ 100 (varsayılan 70) |
| Ortalama katkı | Kategorinin toplam katkısı ÷ toplam satış adedi |
| Maliyet oranı % | Porsiyon maliyeti ÷ satış fiyatı (KDV hariç) × 100 |

| Sınıf | Popüler (karışım ≥ eşik) | Katkısı yüksek (katkı ≥ ortalama) | Önerilen yön |
|---|---|---|---|
| Yıldız | Evet | Evet | Koruyun |
| At | Evet | Hayır | Kârlılığı artırın |
| Bulmaca | Hayır | Evet | Satışı artırın |
| Köpek | Hayır | Hayır | Çıkarmayı veya yeniden tasarlamayı değerlendirin |

Kalemler kendi kategorisi içinde karşılaştırılır; ana yemek ile tatlı kıyaslanmaz.

**Örnek:** 5 kalemli kategori, toplam 1.050 adet. Popülerlik eşiği (100 ÷ 5) × %70 = %14; ortalama katkı 263,81.
Kuzu Tandır 100 adet (%9,52), katkı 420 → popüler değil, katkısı yüksek → **Bulmaca**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --satislar menu_satislari.xlsx
python main.py --satislar menu_satislari.xlsx --kdv 10 --esik 70
python main.py --satislar menu_satislari.xlsx --tum-menu
```

Sütunlar: Menü Kalemi, Satış Adedi, Satış Fiyatı, Porsiyon Maliyeti (zorunlu); Kategori. Fiyatlar KDV dahilse `--kdv`
ile oranı verin; verilmezse fiyatlar KDV hariç kabul edilir.

## Çıktı

`cikti/menu_muhendisligi.xlsx`:

- `Analiz`: kalem bazında sınıf, menü karışımı, katkı payı, toplam katkı, maliyet oranı ve önerilen aksiyon.
- `Kategori Özeti`: ciro, toplam ve ortalama katkı, popülerlik eşiği, sınıf sayıları.
- `Uyarılar`: maliyeti fiyatını aşan kalemler, tek kalemli veya satışı olmayan kategoriler, okunamayan satırlar.

## Dikkat

- **Sınıflandırma görelidir.** Her kalem kategorisinin ortalamasıyla kıyaslanır; her kategoride mutlaka ortalamanın
  altında kalemler çıkar. "Köpek" kötü ürün değil, o listede göreli olarak zayıf üründür.
- **Dönem ve kapsam önemlidir.** Sezonluk ürünler, yeni eklenen kalemler ve stokta bulunmadığı günler satış adedini
  düşürür; analiz dönemini buna göre seçin.
- **Katkı payı yalnız yiyecek maliyetini düşer.** İşçilik, hazırlık süresi ve enerji dahil değildir; emeği yoğun bir
  "yıldız" beklenenden az kazandırabilir.
- **KDV oranı tek değer alınır.** Farklı KDV oranlı kalemler (örneğin alkollü içecek) varsa fiyatları KDV hariç girin.
- **Aksiyonlar genel önerilerdir.** İmza yemekler, çocuk menüsü veya müşteri çeken düşük kârlı ürünler menüde
  bilinçli olarak tutulabilir; karar sizindir.
- **Örnek veri:** Kalemler, adetler, fiyatlar ve maliyetler kurgusaldır.

## Testler

- Dört sınıfın hesabı, menü karışımı, ağırlıklı ortalama katkı, popülerlik eşiği ve sınır eşitlikleri.
- `--esik`, `--kdv` ve `--tum-menu` etkisi; birden çok kategori, tek kalemli ve satışı olmayan kategori.
- Maliyeti fiyatını aşan kalem, okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
