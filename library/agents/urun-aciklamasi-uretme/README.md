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

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.
- **Yapay zekâ hata yapabilir.** Model yanlış, eksik veya uydurma bilgi üretebilir. Kod kontrolleri bazı hataları
  yakalar, hepsini yakalayamaz. Her taslağı okuyup onaylamadan kullanmayın, göndermeyin, imzalamayın.
- **Maskeleme sınırlıdır.** Yalnız telefon, e-posta, TCKN, IBAN ve LinkedIn adresi kalıpları maskelenir. Ad soyad,
  adres, doğum tarihi, sağlık bilgisi ve serbest metindeki diğer kişisel veriler modele gidebilir. Hangi verinin bulut
  sağlayıcısına gönderildiğinden siz sorumlusunuz; gizli veride yerel model (Ollama) kullanın.
- **API ücretleri, kotalar ve API anahtarının güvenliği** size aittir.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
