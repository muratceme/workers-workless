# Sosyal Medya İçerik Takvimi · AI Agent

> Pazarlama › Dijital Pazarlama Uzmanı · Workers / Workless

Kampanyalara ve özel günlere göre **aylık sosyal medya içerik takvimi** kurar ve her gönderi için metin taslağı,
hashtag, görsel fikri ve eylem çağrısı yazar. Taslakları platform sınırları, yasaklı ifadeler, kanıt gerektiren
iddialar ve anma günlerindeki satış dili açısından denetler. Yayın kararı insandadır.

## Nasıl çalışır?

1. **Kod (takvim):**
   - Platform başına haftalık gönderi sayısına göre günleri seçer.

     | Haftalık gönderi | Günler |
     |---|---|
     | 1 | Çarşamba |
     | 2 | Salı, Perşembe |
     | 3 | Pazartesi, Çarşamba, Cuma |
     | 4 | Pazartesi, Salı, Perşembe, Cuma |
     | 5–7 | Hafta içi her gün (+ Cumartesi, Pazar) |

   - **Özel günler:** Sabit tarihli günler kodda tanımlıdır (ör. 23 Nisan, 10 Kasım, Öğretmenler Günü, Dünya Çevre
     Günü). Anneler Günü (Mayıs'ın 2. Pazarı), Babalar Günü (Haziran'ın 3. Pazarı) ve Kasım indirim cuması
     hesaplanır. Dinî bayramlar ve sektörünüze özel günler `--ozel-gunler` dosyasından eklenir.
     - Özel gün, o gün gönderisi olan platformlarda paylaşılır. O gün planlı gönderi yoksa profildeki ilk platforma
       eklenir.
     - **Anma ve farkındalık günleri** "satış dili yok" olarak işaretlenir (ör. 10 Kasım, 15 Temmuz, 25 Kasım).
   - **Kampanyalar:** Kampanya döneminde ilgili platformdaki her iki gönderiden biri kampanyaya ayrılır. Son 3 gün
     brife "bitişe N gün" notu eklenir.
   - **İçerik sütunları:** Kalan günler profildeki sütunlara sırayla dağıtılır (eğitici, ürün, müşteri deneyimi...).
2. **Model (taslak):** Marka profili ve brife göre her gönderinin metnini, hashtaglerini, görsel fikrini ve eylem
   çağrısını yazar. Brifte olmayan oran, fiyat veya müşteri yorumu uydurmaması istenir.
3. **Kod (denetim):**

   | Kontrol | Önem |
   |---|---|
   | Platform karakter sınırı aşıldı (X 280; Instagram ve TikTok 2.200; LinkedIn 3.000; hashtagler dahil) | Yüksek |
   | Profildeki yasaklı ifade kullanıldı | Yüksek |
   | Anma / farkındalık gününde satış ifadesi veya oran / tutar | Yüksek |
   | Model gönderi için taslak döndürmedi | Yüksek |
   | Kanıt gerektiren iddia: "en iyi", "en ucuz", "bir numara", "lider", "garanti", "%100", sağlık beyanı... | Orta |
   | Metinde kampanya mesajında veya profilde olmayan oran / TL tutarı | Orta |
   | Hashtag sayısı platform için önerilenden fazla (X 2, LinkedIn 5, Instagram 8) | Bilgi |

`--model-yok` ile API kullanılmaz; yalnız takvim iskeleti ve brifler çıkar.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                   # örnek: Kasım 2026, Instagram + LinkedIn + X
python agent.py --ay 2026-12 --profil marka_profili.txt --kampanyalar kampanyalar.xlsx --ozel-gunler ozel_gunler.csv
python agent.py --ay 2026-12 --profil marka_profili.txt --model-yok
```

Marka profili (`Alan: değer` satırları; listeler `;` ile):

```text
Marka: Örnek Kahve Atölyesi
Sektör: Nitelikli kahve; çevrim içi satış ve iki mağaza
Hedef Kitle: 25-45 yaş, evde kahve demleyenler
Ton: Samimi, bilgili, sade; "siz" hitabı
Platformlar: Instagram=3; LinkedIn=1; X=2
İçerik Sütunları: Demleme ipuçları; Ürün tanıtımı; Müşteri deneyimi
Yasaklı İfadeler: ucuz; son fırsat
Sabit Hashtagler: #ÖrnekKahveAtölyesi
```

| Dosya | Sütunlar |
|---|---|
| Kampanyalar | Başlangıç, Bitiş, Kampanya, Mesaj (oran ve koşullar burada yazılmalı), Ürün, Platformlar (boşsa tümü) |
| Özel günler | Tarih, Ad, Tür (Kutlama / Anma / Farkındalık) |

## Çıktı

- `icerik_takvimi.xlsx`:
  - `Takvim`: tarih, gün, platform, tür, konu, brif, metin taslağı, hashtagler, görsel fikri, eylem çağrısı, karakter
    sayısı, kontrol sonucu; boş "Onay" sütunu.
  - `Özet`, `Özel Günler`, `Uyarılar`.
- `icerik_takvimi.csv`: planlama araçlarına aktarmak için (UTF-8, `;` ayraçlı).

Mor hücreler yapay zekâ taslağıdır.

## Dikkat

- **Her gönderiyi okuyun.** Model güncel olayları bilmez; yas, afet veya gündem nedeniyle planlı içerik uygunsuz
  hâle gelebilir. Otomatik yayına bağlamayın.
- **Reklam mevzuatı:** Üstünlük ve karşılaştırma iddiaları ispatlanabilir olmalıdır; sağlık beyanları, indirim
  duyuruları (indirim öncesi fiyat) ve iş birliği / reklam etiketleri için Ticari Reklam ve Haksız Ticari
  Uygulamalar Yönetmeliği ile Reklam Kurulu kılavuzlarına bakın. Koddaki iddia listesi bir ön elemedir.
- **Platform sınırları değişebilir.** Koddaki karakter ve hashtag sınırlarını platformların güncel kurallarıyla
  karşılaştırın.
- **Özel gün listesi** genel bir başlangıçtır. Dinî bayram tarihleri her yıl değişir; dosyadan ekleyin.
- **Örnek veri:** Örnek marka ve kampanyalar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Anneler / Babalar Günü ve Kasım indirim cuması tarihleri; dosyadan özel gün.
- Takvim iskeleti: platform başına gönderi sayısı, anma günü işareti, kampanyanın platform kısıtı.
- Denetimler: iddia, karakter sınırı, yasaklı ifade, hashtag sayısı, anma gününde satış, doğrulanmamış oran,
  eksik taslak; hashtag temizliği.
- Excel, CSV ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
