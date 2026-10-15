# Misafir Yorum Analizi · AI Agent

> Otel › Misafir İlişkileri · Workers / Workless

Booking, Google, TripAdvisor gibi platformlardaki **çok dilli misafir yorumlarını** konu ve duygu bazında analiz
eder. Yönetimin hemen ilgilenmesi gereken **acil** yorumları ayırır ve haftalık rapor çıkarır. Her yorum için
**yorumun kendi dilinde** cevap taslağı hazırlar. Hiçbir cevap otomatik yayımlanmaz.

## Nasıl çalışır?

1. **Kod:**
   - Platform puanlarını 10'luk ölçeğe çevirir (Booking 10, Google ve TripAdvisor 5, HolidayCheck 6;
     `--olcek` ile değiştirilir).
   - Haftalık ortalamaları ve platform dağılımını hesaplar.
2. **Model:** Her yorumu Türkçe özetler ve 13 konuda (temizlik, oda, yemek, personel, konum, fiyat/değer,
   havuz-plaj, gürültü, Wi-Fi, check-in/out, animasyon, genel deneyim, diğer) duygu çıkarır. Acil durumları
   işaretler: hijyen ve gıda güvenliği, sağlık, güvenlik, ayrımcılık, kaba davranış.
3. **Kod:** Sonuçları toplar ve şunları çıkarır:
   - Konu bazında bahsedilme sayısı ve net duygu skoru (−100…+100)
   - Önceki haftaya göre olumsuz bahsedilmedeki değişim
   - **Puanla yorumun çeliştiği** kayıtlar (ör. 1/5 puan ama yorum olumlu)

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek 10 yorum (TR, EN, DE, RU)
python agent.py --girdi yorumlar.xlsx
python agent.py --girdi yorumlar.xlsx --olcek Booking=10 Google=5 Tatilsepeti=10 --cevap-yok
```

## Çıktı

`Acil` (yapılacak/sorumlu sütunuyla) · `Konu Analizi` (grafikli, haftalık) · `Yorumlar` (özet, konular,
alıntılar, kontrol, cevap taslağı, "Onay")

## Testler

Testler gerçek API çağırmaz.

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
