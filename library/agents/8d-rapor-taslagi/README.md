# 8D Rapor Taslağı · AI Agent

> Üretim › Kalite · Workers / Workless

Müşteri şikâyeti veya uygunsuzluk kaydından **8D raporunun D1–D8 adımlarını içeren taslağı** hazırlar.
Muayene verisi varsa önce **kodla analiz eder**: lot bazında PPM, etkilenen lotlar, makine/kalıp kırılımı, hata
Pareto'su ve şüpheli üretim aralığı. Taslak 8D ekibinin üzerinde çalışacağı bir başlangıç belgesidir.

## Nasıl çalışır?

1. **Veri analizi (kod):** Muayene kayıtları lot bazında toplanır. Etkilenen lotlar şikâyette geçen lotlar ile
   ortalamanın 2 katından yüksek PPM'li lotlardır.
2. **Taslak (model):** Şikâyet metni (telefon ve e-posta maskeli), parça/ekip bilgisi ve kod bulguları gönderilir.
   Model şunları yazar:
   - **D2:** 5N1K ve Is / Is Not karşılaştırması
   - **D3:** Geçici önlemler (ayıklama kapsamı, temiz parça etiketi)
   - **D4:** **Oluşum** ve **kaçış** kök nedenleri ayrı ayrı; 6M sınıfı ve 5 Neden zinciriyle. Veriyle
     desteklenmeyen her neden **"doğrulanmalı"** hipotezi olarak işaretlenir ve doğrulama yöntemi yazılır.
   - **D5–D6:** Kalıcı faaliyetler. Her biri D4'teki kök neden numarasına (K1, K2…) bağlanır.
   - **D7:** Önleme (PFMEA, kontrol planı, yatay yayılım)
   - **D8:** Kapanış
   - Toplanacak veri listesi
3. **Denetim (kod):** Taslakta şunlar aranır ve eksikler raporda işaretlenir:
   - Oluşum ve kaçış kök nedeninin ikisi de var mı?
   - Her kök nedenin kalıcı faaliyeti var mı?
   - Olmayan bir kök nedene bağlanmış faaliyet var mı?

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                         # örnek şikâyet + muayene + bilgi
python agent.py --sikayet sikayet.txt --muayene muayene.xlsx --bilgi bilgi.json
python agent.py --sikayet uygunsuzluk.docx                               # yalnız metinle
```

- **Muayene kayıtları:** Tarih, Lot, Makine, Kalıp, Kontrol Edilen, Hatalı, Hata Türü. Aynı lotta birden çok
  hata türü ayrı satırda olabilir.
- **bilgi.json:** müşteri, parça no ve adı, şikâyet no, tespit tarihi, ekip (rol ve görev), hedef süreler.

## Çıktı

- **Excel:** `Özet` · `D1 Ekip` · `D2 Problem` · `D3 Geçici Önlemler` · `D4 Kök Neden` · `D5-D6 Kalıcı
  Faaliyetler` · `D7 Önleme` · `Veri İhtiyacı` · `Veri Analizi` (lot PPM grafiği). Her sayfada "Onay" sütunu
  vardır.
- **Markdown:** Aynı içerikte `.md` taslak.

> "doğrulanmalı" işaretli kök nedenler hipotezdir. Veri toplanıp doğrulanmadan müşteriye kesin kök neden
> olarak bildirilmemelidir.

## Testler

Testler gerçek API çağırmaz.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
