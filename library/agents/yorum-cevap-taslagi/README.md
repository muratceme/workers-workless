# Yorum Cevap Taslağı · AI Agent

> Misafir İlişkileri › Misafir İlişkileri Müdürü · Workers / Workless

Google, Booking, TripAdvisor gibi platformlardaki misafir yorumlarına otelin **üslup rehberine** uygun, kişiselleştirilmiş ve
yorumun dilinde cevap taslağı hazırlar. Herkese açık cevapta yasak olanları (oda numarası, tazminat teklifi, suçlama, kalıp
cevap) kodla denetler. Cevaplar yayımlanmaz.

## Nasıl çalışır?

1. **Hazırlık (kod):** Üslup rehberi ve imza satırı ("İmza: …") okunur. Misafir adı modele gönderilmez; cevapta `{AD}` yer
   tutucusu kullanılır, kod ilk adla doldurur.
2. **Taslak (model):**
   - Yorumdaki olumlu ve olumsuz noktalar alıntıyla çıkarılır.
   - Yorumun dilinde, somut ayrıntılara değinen bir cevap yazılır.
   - Yönetim için yayımlanmayan bir iç not eklenir (ör. kontrol edilmesi gereken oda).
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o cevap modele geri gönderilir (varsayılan en fazla 2 tur):
   - Alıntı yorumda yok
   - Olumsuz bir nokta cevapta ele alınmamış
   - Tazminat, indirim, iade ya da ücretsiz konaklama teklifi (`--teklif-izni` verilmedikçe)
   - Misafiri suçlayan ifade
   - Oda ya da rezervasyon numarası olabilecek 3 ve üstü haneli sayı ya da yorumda olmayan sayı
   - İmza ya da `{AD}` hitabı yok
   - `--azami` karakter aşılmış
   - Başka bir cevapla neredeyse aynı (kalıp cevap)

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                        # örnek yorumlar + üslup rehberi
python agent.py --girdi yorumlar.xlsx --rehber uslup_rehberi.txt
python agent.py --girdi yorumlar.xlsx --rehber uslup_rehberi.txt --azami 1000 --tur 3
```

Üslup rehberine ton, hitap şekli, yasaklar, misafirin yönlendirileceği iletişim adresi ve "İmza: …" satırını yazın.

## Çıktı

- `cikti/yorum_cevaplari.xlsx`: `Cevaplar` sayfası. Olumlu ve olumsuz noktalar, cevap taslağı, karakter sayısı, iç not ve
  "Onay" sütunu.

## Dikkat

- **Cevaplar yayımlanmaz.** Okuyup onayladıktan sonra platformda yayımlayın. Olumsuz noktaları ilgili departmana iletin.
- **Platform kuralları.** Platformların cevap uzunluğu ve içerik kuralları farklıdır; `--azami` değerini buna göre ayarlayın.
- **Çeviri kontrolü.** Bilmediğiniz dillerdeki cevapları yayımlamadan önce kontrol ettirin.
- **Veri gönderimi.** Yorum metni ve üslup rehberi seçtiğiniz yapay zekâ sağlayıcısına gönderilir; misafir adı gönderilmez.
  Yorumlar zaten herkese açıktır, ancak telefon ve e-posta yine de maskelenir.
- **Örnek veri** kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek cevapta misafirin oda numarasını yazar, indirim teklif eder ve
bir olumsuz noktaya (klima) değinmez. Kodun bunları yakalayıp yalnız bu cevapları yeniden yazdırdığı doğrulanır; kalıp
cevap, eksik imza, suçlayıcı ifade ve uzunluk kontrolleri de test edilir.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve yayımlanan cevapların doğruluğu kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
