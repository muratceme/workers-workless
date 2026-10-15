# Ürün Bilgi Kartı Özeti · AI Agent

> Mağaza › Satış Danışmanı · Workers / Workless

Ürün teknik bilgilerinden **müşteriye mağazada anlatılacak kısa ürün kartları** hazırlar: 3 öne çıkan fayda, kısa
anlatım, aynı kategorideki ürünlerle fark özeti ve sık sorulan sorular. Her özellik değeri ve karşılaştırma ürün
tablosuyla birebir kontrol edilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Ürün tablosu okunur. SKU, Ürün, Kategori ve Fiyat dışındaki her sütun bir teknik özelliktir; boş
   hücre "bilgi yok" demektir.
2. **Kart (model):** Her ürün için öne çıkanlar (özellik + değer + müşteriye fayda), kısa anlatım, aynı kategorideki
   ürünlerle farklar ve sık sorulanlar yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o ürünün kartı modele geri gönderilir (varsayılan en fazla
   2 tur):
   - Adı geçen özellik tabloda yok ya da yazılan değer tablodakiyle farklı ("8 saat" ↔ tabloda "6 saat")
   - Karşılaştırılan ürün başka kategoride
   - Metinde ürünün özelliklerinde ya da fiyatında olmayan bir sayı var
   - "Su geçirmez", "en iyi", "en ucuz", "ömür boyu", "kırılmaz" gibi iddia özelliklerde yazmıyor (ör. IPX4 su geçirmez
     değildir)
   - 3 öne çıkan fayda yazılmamış

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                 # örnek kulaklık ve süpürgeler
python agent.py --girdi urunler.xlsx
python agent.py --girdi urunler.xlsx --kategori "Kablosuz Kulaklık"
```

Tablo örneği:

| SKU | Ürün | Kategori | Fiyat | Pil Süresi | Suya Dayanıklılık | Ağırlık |
|---|---|---|---|---|---|---|
| K-101 | Örnek Kulaklık Air | Kablosuz Kulaklık | 1.899 TL | 6 saat (kutuyla 24 saat) | IPX4 | 4,5 g |

## Çıktı

- `cikti/urun_kartlari.xlsx`:
  - `Kartlar`: öne çıkanlar, kısa anlatım, sık sorulanlar, kontrol sonucu.
  - `Farklar`: aynı kategorideki ürünlerle özellik karşılaştırması.
- `cikti/urun_kartlari.txt`: kontrolden geçen kartların mağazada kullanılacak kısa hâli.

## Dikkat

- **Kaynak sizin tablonuz.** Agent internetten bilgi çekmez; kartlar tablonuz kadar doğrudur. Teknik değerleri
  üretici belgeleriyle, fiyatları güncel etiketle karşılaştırın.
- **Tüketici bilgilendirmesi.** Müşteriye verilen bilgi yanıltıcı olmamalıdır. Kod belirgin iddiaları yakalar; son kontrol
  sizdedir.
- **Örnek veri:** Ürünler, markalar ve fiyatlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk kartlarda bilerek:
- Air kulaklığın pil süresini 8 saat yazar ve "tamamen su geçirmez" der,
- Pro'yu Sport ile karşılaştırırken Sport'un değerini IP68 yazar,
- süpürgenin pil süresini 45 dakika yazar ve onu bir kulaklıkla karşılaştırır.

Kodun bunları yakaladığı ve yalnız sorunlu kartları yeniden yazdırdığı doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve müşteriye verilen bilgilerin doğruluğu kullanıcının
sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
