# Sepet Birliktelik Analizi · Kod Bloğu

> CRM ve Sadakat Programı › CRM Uzmanı · Workers / Workless

Fiş satırlarından **birlikte alınan ürün çiftlerini destek, güven ve lift** değerleriyle çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül | Anlamı |
|---|---|---|
| Destek % | Birlikte fiş ÷ toplam fiş × 100 | Çift, fişlerin yüzde kaçında var |
| Güven A→B % | Birlikte fiş ÷ A'nın fiş sayısı × 100 | A alanların yüzde kaçı B de aldı |
| Lift | (Birlikte fiş × toplam fiş) ÷ (A fiş × B fiş) | 1'den büyük: beklenenden sık; 1: ilişki yok; 1'den küçük: beklenenden seyrek |

- Bir fişte aynı ürün birden çok satırda geçse de bir kez sayılır; adet ve tutar hesaba girmez.
- Tek ürünlü fişler toplam fiş sayısına dahildir.
- Ürün adları büyük-küçük harf ve Türkçe karakter ayırmadan eşleştirilir ("ÇAY" = "çay").
- Çiftler lift değerine göre büyükten küçüğe sıralanır.

**Örnek:** 10 fişin 5'inde çay, 4'ünde şeker, 3'ünde ikisi birlikte → destek %30; güven çay→şeker 3 ÷ 5 = %60,
şeker→çay 3 ÷ 4 = %75; lift (3 × 10) ÷ (5 × 4) = 1,50.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (en az 2 fiş)
python main.py --fisler fis_satirlari.xlsx --min-fis 20 --min-destek 0,5 --min-lift 1
```

Sütunlar: Fiş No (veya Sipariş No), Ürün. Her satır bir fiş kalemidir.

| Parametre | Varsayılan | Anlamı |
|---|---|---|
| `--min-fis` | 5 | Çiftin birlikte geçmesi gereken en az fiş sayısı |
| `--min-destek` | 0 | En az destek % |
| `--min-lift` | 0 | En az lift (yalnız olumlu ilişkiler için 1) |

## Çıktı

`cikti/sepet_birliktelik.xlsx`:

- `Ürün Çiftleri`: çift, birlikte fiş sayısı, destek, iki yönlü güven, lift; boş "Aksiyon" sütunu.
- `Ürünler`: her ürünün geçtiği fiş sayısı ve desteği.
- `Özet`: fiş sayısı, tek ürünlü fiş, ortalama sepet çeşidi.
- `Uyarılar`: eksik satırlar, az fiş uyarısı.

## Dikkat

- **Birliktelik nedensellik değildir.** İki ürünün birlikte alınması, birinin diğerini sattırdığını göstermez; ikisi
  de aynı kampanyada veya aynı reyonda olabilir.
- **Az veride lift yanıltır.** Seyrek satan iki ürün birkaç fişte birlikte geçince lift çok yüksek çıkar. `--min-fis`
  değerini veri büyüklüğünüze göre yükseltin; 100'den az fişte uyarı verilir.
- **Çok satan ürün her şeyle birlikte görünür.** Ekmek, poşet gibi ürünlerde destek yüksek, lift 1'e yakındır; kararı
  lift ile birlikte verin.
- **Yalnız ikili çiftler.** Üçlü ve daha büyük ürün kümeleri hesaplanmaz.
- **Ürün düzeyi önemlidir.** SKU bazında sonuçlar dağınık kalıyorsa ürün sütununa kategori veya alt grup yazıp
  yeniden çalıştırın.
- **Kişisel veri:** Dosyada müşteri adı, kart numarası gibi alanlara gerek yoktur; yalnız fiş no ve ürün yeter.
- **Örnek veri:** Fişler ve ürünler kurgusaldır.

## Testler

- Destek, iki yönlü güven ve lift; lift 1 ve 1'den küçük durumlar.
- Eşikler (en az fiş, destek, lift) ve sıralama; üç ürünlü sepetten üç çift.
- Yazım farklarının birleştirilmesi, eksik satırlar ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
