# Kanallar Arası Stok ve Fiyat Kontrolü · Kod Bloğu

> Pazaryeri Yönetimi › Pazaryeri Uzmanı · Workers / Workless

Farklı satış kanallarındaki **stok ve fiyatları merkezi listeyle karşılaştırıp uyumsuzlukları** listeler. Kanal
panellerine bağlanmaz; panellerden indirdiğiniz listelerle çalışır.

## Nasıl hesaplar?

| Sorun | Koşul | Önem |
|---|---|---|
| Merkezde yok | Kanaldaki stok kodu merkezi listede bulunmuyor | Yüksek |
| Stok yokken satışta | Merkez stok 0, kanal stok > 0 | Yüksek |
| Fiyat düşük | Kanal fiyatı beklenen fiyatın `--fiyat-tolerans` (%1) üzerinde altında | Yüksek |
| Kanal stoku merkezden fazla | Kanal stok − merkez stok > `--stok-tolerans` (0) | Orta |
| Kanalda stok kapalı | Merkez stok > 0, kanal stok 0 | Orta |
| Fiyat yüksek | Kanal fiyatı beklenen fiyatın tolerans üzerinde üstünde | Orta |
| Kanal stoku eksik | Merkez stok − kanal stok > tolerans; yalnız `--tam-stok` verilirse | Bilgi |

- **Beklenen fiyat** = merkez fiyat × (1 + kanal fiyat farkı %). Kanalda bilerek farklı fiyat uyguluyorsanız (ör.
  komisyon nedeniyle +%10) kanal ayar dosyasına yazın; verilmezse fark 0 alınır.
- **Fiyat farkı %** = (kanal fiyatı − beklenen fiyat) ÷ beklenen fiyat × 100. Tam tolerans sınırındaki fark uyumludur.
- **Kanalda olmayan ürünler:** merkezde stoku olan ama o kanalda hiç listelenmemiş stok kodları ayrı sayfada verilir.
- Stok kodları büyük-küçük harf ayırmadan eşleştirilir.

**Örnek:** merkez fiyat 2.000, kanal farkı %10 → beklenen 2.200; kanalda 2.000 → fark −%9,09 → "Fiyat düşük".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --merkez merkez_liste.xlsx --kanal kanal_listesi.xlsx --kanal-ayar kanal_ayarlari.xlsx
python main.py --merkez m.xlsx --kanal k.xlsx --fiyat-tolerans 2 --stok-tolerans 1 --tam-stok
```

| Dosya | Sütunlar |
|---|---|
| Merkezi liste | Stok Kodu, Stok (zorunlu); Ürün, Fiyat |
| Kanal listesi | Kanal, Stok Kodu (zorunlu); Stok, Fiyat. Tüm kanallar alt alta, "Kanal" sütununda kanal adı |
| Kanal ayarları (isteğe bağlı) | Kanal, Fiyat Farkı % |

## Çıktı

`cikti/kanal_stok_fiyat_kontrolu.xlsx`:

- `Uyumsuzluklar`: önem sırasıyla sorunlu kayıtlar, stok ve fiyat farkları; boş "Yapılan" sütunu.
- `Kanal Özeti`: kanal başına kayıt, uyumlu ve sorun türü sayıları.
- `Kanalda Olmayan Ürünler`: stokta olup kanalda listelenmemiş ürünler.
- `Tüm Kayıtlar`: uyumlu olanlar dahil her kayıt.
- `Uyarılar`: mükerrer kodlar, okunamayan değerler.

## Dikkat

- **Listeler aynı anda alınmış olmalı.** Stok dakikalar içinde değişir; farklı saatlerde indirilen listelerde farkın
  bir kısmı zamanlamadan kaynaklanır. Küçük farkları `--stok-tolerans` ile görmezden gelebilirsiniz.
- **Stok kodu ortak olmalı.** Kanalda farklı kod (kanal barkodu, satıcı stok kodu) kullanılıyorsa önce eşleştirme
  tablosuyla merkez koduna çevirin; yoksa tüm kayıtlar "Merkezde yok" çıkar.
- **Fiyat esası aynı olmalı.** Merkez ve kanal fiyatları aynı KDV esasında ve aynı para biriminde olmalıdır. Kanaldaki
  kampanya, kupon ve sepet indirimleri liste fiyatında görünmez.
- **Stok tamponu kullanıyorsanız** kanal stokunun merkezden az olması normaldir; bu yüzden "Kanal stoku eksik" yalnız
  `--tam-stok` ile raporlanır.
- **Düzeltme yapmaz.** Kod yalnız listeler; kanal panelinde değişiklik yapmaz.
- **Örnek veri:** Kanallar, ürünler ve fiyatlar kurgusaldır.

## Testler

- Stok kuralları ve tolerans sınırları; fiyat toleransı (tam %1), kanal fiyat farkı, eksik fiyat.
- Merkezde olmayan kod, kanalda olmayan ürün, kanal özeti, ayarı olup kaydı olmayan kanal.
- Mükerrer ve okunamayan satırlar; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
