# Ürün Değerlendirme Analizi · AI Agent

> Müşteri Deneyimi › Müşteri Hizmetleri Temsilcisi · Workers / Workless

Ürün yorumlarını **duygu ve konu** bazında sınıflandırır: ürün kalitesi, beden / kalıp, kargo, paketleme, fiyat /
performans, açıklamaya uygunluk, satıcı ilgisi ve ürün güvenliği. Ürün bazında konu skorlarını ve örnek alıntıları çıkarır;
ürün güvenliği yorumlarını ayrı bir "Acil" sayfasında toplar.

## Nasıl çalışır?

1. **Ön inceleme (kod):** Yorumlarda ürün güvenliği belirtileri işaretlenir: duman, yanma, alev, patlama, elektrik çarpması,
   kızarıklık, kaşıntı, alerji, zehirlenme, kesik. Kargo ve paketleme ifadeleri de işaretlenir.
2. **Sınıflandırma (model):** Her yorum için genel duygu, konular (her biri duygu ve alıntıyla) ve tek cümlelik Türkçe özet
   yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız ilgili yorum modele geri gönderilir (varsayılan en fazla 2 tur):
   - Yorumda olmayan alıntı
   - Güvenlik belirtisi olan yorumda "Ürün güvenliği" konusu yok
   - Kargo ya da paketleme ifadesi olan yorumda ilgili konu yok
4. **Özet (kod):**
   - Ürün bazında yorum sayısı ve ortalama puan
   - Konu bazında olumlu / olumsuz / net skor ve en çok şikâyet edilen konu
   - Puanla duygusu çelişen yorumlar (ör. 5 yıldız ama olumsuz metin) işaretlenir

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                              # örnek yorumlar
python agent.py --girdi yorumlar.xlsx
python agent.py --girdi yorumlar.xlsx --tur 3
```

## Çıktı

- `cikti/urun_degerlendirme.xlsx`:
  - `Acil`: ürün güvenliği yorumları ve yapılacaklar.
  - `Ürün Özeti`: konu bazında skorlar ve örnek alıntılar.
  - `Yorumlar`: duygu, konular, özet ve puan-duygu çelişkisi.

## Dikkat

- **Ürün güvenliği yorumlarını mutlaka okuyun.** Anahtar kelime taraması farklı ifade edilmiş olayları kaçırabilir. Güvenlik
  şikâyetlerinde müşteriyle iletişim ve ürünün kalite / ürün güvenliği birimine bildirimi şirket prosedürünüze göre
  yapılmalıdır.
- **Etiketler yapay zekâ sınıflandırmasıdır.** Örnek alıntılarla birlikte okuyun; ironi ve kısa yorumlar yanlış
  sınıflandırılabilir.
- **Veri gönderimi.** Yorum metni, ürün adı ve puan seçtiğiniz yapay zekâ sağlayıcısına gönderilir. Telefon ve e-posta
  maskelenir; kullanıcı adları gönderilmez.
- **Örnek veri** kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek fişten duman çıkan ve ciltte kaşıntı yapan yorumları "ürün
kalitesi" sayar ve bir alıntıyı değiştirir. Kodun bunları yakalayıp yalnız bu yorumları yeniden sorduğu, ürün özetini ve
"Acil" sayfasını doğru oluşturduğu doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve analizin doğruluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
