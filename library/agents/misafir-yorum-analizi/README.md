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

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
