# Reklam Metni Varyasyonları · AI Agent

> E-Ticaret › Performans Pazarlama Uzmanı · Workers / Workless

Ürün bilgisinden arama ağı reklamları (Google Ads **duyarlı arama reklamı**) için şunları yazar:
- 3–15 başlık,
- 2–4 açıklama,
- 2 görünen URL yolu.

**Karakter sınırlarını, iddiaları ve tekrarları kod denetler.** Sonuç Google Ads Editor'e kopyalanabilir sütun
düzenindedir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Ürün dosyası, kanal limitleri ve kaçınılacak ifadeler okunur.
2. **Yazım (model):** Farklı açılardan (ürün adı, malzeme, kullanım, fiyat, kampanya, çağrı) başlık ve açıklamalar
   yazılır; yalnız ürün bilgisindeki iddialar kullanılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o ürün modele geri gönderilir (varsayılan en fazla 2 tur):
   - İstenen adette başlık / açıklama yok; boş metin
   - Karakter sınırı aşıldı (boşluk ve noktalama dahil)
   - Tekrar eden başlık
   - Ürün bilgisinde olmayan sayı (fiyat, indirim oranı, renk sayısı …)
   - İndirim / kampanya alanı boşken "indirim, kampanya, fırsat, sezon sonu" dili
   - Kaçınılacak ifade ("en iyi", "garantili", "%100" …). Ürün bilgisinde aynen geçen ifadeye (ör. "%100 pamuk") izin
     verilir.
   - Anahtar kelime `--anahtar-min` (varsayılan 3) başlıktan azında geçiyor
   - Tamamı büyük harf metin; yol alanında Türkçe karakter ya da boşluk
4. **Çıktı (kod):** Karakter sayıları ve kalan karakter hesaplanır; sınırı aşan hücre kırmızı.

## Karakter sınırları

`ornek_veri/kanal_limitleri.csv` dosyasında tutulur ve düzenlenebilir:

| Alan | Sınır |
|---|---|
| Başlık | 30 |
| Açıklama | 90 |
| Yol | 15 |

Kaynak: Google Ads Yardım, "Duyarlı arama reklamları hakkında" (reklam başına en fazla 15 başlık ve 4 açıklama).
Google'ın kendisi metni kısaltabilir ya da başlıkları farklı sırayla birleştirebilir. Başka bir kanal için dosyaya kendi
sınırlarınızı yazın.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 2 ürün, 10 başlık + 4 açıklama
python agent.py --urunler urunler.xlsx --baslik 15 --aciklama 4 --anahtar-min 4
```

## Çıktı

`cikti/reklam_metinleri.xlsx`:
- `Reklamlar`: ürün başına Headline 1…N, Description 1…4, Path 1–2, son kontrol (mor: model, kırmızı: sınır aşımı).
- `Karakter Sayıları`: her metnin uzunluğu ve kalan karakter.
- `Kontrol`

## Dikkat

- **Reklam politikası ve mevzuat** sizin sorumluluğunuzdadır: Google Ads politikaları, ticari reklam ve indirimli satış
  kuralları, marka hakları. Kaçınılacak ifadeler listesi yalnız bir başlangıçtır; kendi listenizle genişletin.
- **Fiyat ve indirim güncel olmalı.** Metin üretildikten sonra fiyat değiştiyse metinleri yeniden üretin.
- **Performansı test belirler.** Varyasyonları yayınlayıp reklam gücü ve dönüşüm verisine göre eleyin.
- **Örnek veri:** Ürünler ve fiyatlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda indirimi olmayan SW-02'ye "%20 indirim" yazar. Kodun hem olmayan
sayıyı hem indirim dilini yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- adet, karakter sınırı, tekrar, kaçınılacak ifade ve "%100 pamuk" istisnası,
- büyük harf, yol biçimi ve anahtar kelime kuralları,
- çıktı sütun düzeni.

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
