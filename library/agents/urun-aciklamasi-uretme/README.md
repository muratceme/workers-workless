# Ürün Açıklaması Üretme · AI Agent

> E-ticaret › Kategori ve İçerik Yönetimi › E-ticaret İçerik Uzmanı · Workers / Workless

Ürün özellik tablosundan her satış kanalı (web sitesi, pazaryerleri) için şunları üretir:
- Başlık ve meta açıklama
- Ürün açıklaması
- Özellik maddeleri
- Anahtar kelimeler

Üretilen her metni **kodla denetler** ve kural dışı olanları düzeltmesi için modele geri gönderir. Hiçbir metin
otomatik yayımlanmaz; "Onay" sütunu vardır.

## Agent döngüsü

1. **Üret:** Model, kanalın karakter sınırlarına, başlık sırasına ve tonuna göre taslağı yazar. Yalnız tabloda
   verilen özellikleri kullanır; boş özellikten hiç bahsetmez.
2. **Denetle (kod):**
   - Başlık, meta ve açıklama karakter sınırları, özellik maddesi sayısı
   - Başlığın markayla başlaması (kanal kuralıysa)
   - **Ürün verisinde olmayan sayılar** (uydurma ölçü, gramaj, iplik sıklığı vb.)
   - **Kanıtlanamayan üstünlük ve sağlık iddiaları** ("en iyi", "lider", "antibakteriyel", "organik"...).
     Bu ifadeler ürün verisinde açıkça geçiyorsa serbesttir.
3. **Düzelt:** Hatalı taslaklar hata listesiyle modele geri gönderilir (varsayılan 2 tur, `--tur`). Kalan
   hatalar raporda kırmızıyla gösterilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                  # örnek ürünlerle (web sitesi + pazaryeri)
python agent.py --girdi urunler.xlsx --kanallar kanallar.json
python agent.py --girdi urunler.xlsx --kanallar kanallar.json --yasakli yasakli_ifadeler.txt --tur 3
```

- **Ürün tablosu:** `Stok Kodu` zorunludur. Diğer tüm sütunlar (Marka, Ürün Tipi, Malzeme, Renk, Ölçü, Yıkama...)
  özellik olarak kullanılır.
- **kanallar.json:** Her kanal için `baslik_max`, `meta_aciklama_max` (0 = yok), `aciklama_min` / `aciklama_max`,
  `ozellik_maddesi`, `baslik_sirasi` ve `ton` alanları verilir. Örnek dosyadaki sınırlar **örnektir**;
  pazaryerlerinin güncel içerik kurallarını girin.

## Çıktı

`İçerikler` (kanal bazında metinler, karakter sayıları, düzeltme turu, kontrol sonucu) · `Ürün Verisi` · `Bilgi`

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek hata yapar ("300 TC", "en iyi", uzun başlık).
Testler hatanın kodla yakalandığını ve düzeltme turunda giderildiğini doğrular.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri kullanıcıya aittir; yayımlanan içeriklerin doğruluğu ve
reklam mevzuatına uygunluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
