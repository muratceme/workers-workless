# GTİP Sınıflandırma Önerisi · AI Agent

> Dış Ticaret › Dış Ticaret Uzmanı · Gerisi Bende

Ürün tanımlarından **olası GTİP kodlarını gerekçesiyle önerir**. Model yalnız **sizin verdiğiniz tarife tablosundaki**
kodlar arasından seçer; tabloda olmayan bir kod yazarsa, alt pozisyonu olan bir pozisyonda kalırsa ya da ürün bilgisinde
olmayan bir özelliğe dayanırsa öneri modele geri gönderilir. **Kesin sınıflandırma için gümrük müşaviri onayı gerekir.**

## Nasıl çalışır?

1. **Hazırlık (kod):** Ürün listesi ve tarife tablosu okunur. Malzemesi ya da kullanım amacı boş ürünler işaretlenir.
2. **Öneri (model):** Her ürün için 1–3 kod; uygulanan Genel Yorum Kuralları (GYK 1, 3(a)/(b)/(c), 6 …), dayanılan ürün
   özellikleri, gerekçe, güven (Yüksek / Orta / Düşük) ve sınıflandırma için eksik bilgiler.
3. **Denetim (kod):** Aşağıdakilerden biri varsa öneri modele geri gönderilir (varsayılan en fazla 2 tur):
   - Kod tarife tablosunda yok (kod uydurulamaz) ya da alt pozisyonu olan bir pozisyon seçilmiş
   - GYK 1 yok; 4 haneden uzun kodda GYK 6 yok
   - "Dayanılan özellik" ürün bilgisinde aynen geçmiyor; gerekçede ürün bilgisinde / tabloda olmayan sayı
   - Malzeme ya da kullanım amacı eksikken soru sorulmamış ya da "Yüksek" güven verilmiş; birden fazla "Yüksek" öneri
4. **Rapor (kod):** Her öneriye tarife tanımı ve üst pozisyon tanımı eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                    # örnek ambalaj ürünleri (39. ve 48. fasıl, 6 hane)
python agent.py --urunler urunler.xlsx --tarife tgtc_fasil39.xlsx
```

**Tarife tablosu:** Türk Gümrük Tarife Cetveli'nin güncel hâlinden ürünlerinizle ilgili fasılları (12 haneli kodlar ve
tanımlarıyla) "Kod; Tanım" sütunlarıyla kaydedin. Tablo ne kadar ayrıntılıysa öneri o kadar ayrıntılı olur; çok büyük
tablolar maliyeti artırır.

## Çıktı

- `cikti/gtip_onerileri.xlsx`: ürün başına öneriler (güven: yeşil / sarı / kırmızı), eksik bilgi soruları.

## Dikkat

- **Öneridir.** Fasıl ve bölüm notları, açıklama notları ve sınıflandırma kararları bu agent'ta yoktur. Beyan öncesi
  gümrük müşavirine danışın; belirsiz durumlarda Bağlayıcı Tarife Bilgisi (BTB) başvurusu yapın. Yanlış GTİP, vergi farkı
  ve cezaya yol açabilir.
- **Örnek tablo:** `ornek_veri/tarife.csv` Armonize Sistem'in 6 haneli düzeyindeki bazı pozisyonların özet tanımlarıdır,
  resmî tarife metni değildir. Kendi işiniz için güncel resmî tarifeyi kullanın.
- **Örnek veri:** Ürünler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk çıktıda tabloda olmayan 12 haneli bir kod uydurur, şişe için alt pozisyonu
olan 3923'te kalır, üründe olmayan "şeffaf" özelliğine ve "1,5 litre"ye dayanır, malzemesi belirtilmeyen koliye soru
sormadan "Yüksek" güven verir ve bir ürünü atlar. Kodun yedi sorunu da yakaladığı doğrulanır.

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
