# Mülakat Soru Seti Hazırlama · AI Agent

> İnsan Kaynakları › İşe Alım · Workers / Workless

Pozisyon bilgisinden ve yetkinlik listesinden **yapılandırılmış (yetkinlik bazlı) mülakat soru seti**
hazırlar. Mülakatçının dolduracağı **formüllü değerlendirme formunu** da üretir.

## Neler üretir?

- **Sorular:** Her yetkinlik için 2–3 soru: **davranışsal** (STAR ile yanıtlanan geçmiş deneyim), **durumsal**
  (senaryo) veya **teknik**. Her sorunun yanında şunlar yer alır:
  - Takip (sondaj) soruları
  - Dinlenecek **olumlu ve olumsuz göstergeler**
  - Davranışsal **1 – 3 – 5 puan tanımları**
  - Tahmini süre
- **Açılış ve kapanış metinleri:** Mülakatçının okuyacağı kısa metinler.
- **Değerlendirme formu (Excel):** Puan hücreleri sarıdır. Yetkinlik ortalaması ve **ağırlıklı toplam puan**
  formülle hesaplanır. Yalnız puanlanmış yetkinlikler toplama girer.

## Kod denetimi (model taslağının üstünde)

1. **Ayrımcı soru taraması:** Soru ve takip soruları ortak tarayıcıdan geçer (`ayrimcilik.py`). Taranan
   konular: yaş, medeni hâl, çocuk/aile planı, gebelik, sağlık, din, köken/memleket, siyasi görüş, sendika,
   askerlik, görünüş ve gereksiz kişisel veri.
2. **Kapsam:** Her yetkinliğin en az bir sorusu olmalıdır.
3. **Süre:** Toplam süre, açılış ve kapanış dahil (10 dk) mülakat süresini (`--sure`) aşmamalıdır.

Sorun varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur). Kalanlar raporda işaretlenir.
**Dayanak:** 4857 sayılı İş Kanunu md. 5; 6701 sayılı Kanun md. 3 ve 6; 6356 sayılı Kanun md. 25.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                               # örnek: saha satış temsilcisi, 5 yetkinlik
python agent.py --pozisyon ilan.docx --yetkinlikler yetkinlikler.xlsx --sure 60
python agent.py --pozisyon talep_formu.txt --sure 45                           # yetkinlikleri model önersin
```

## Çıktı

`Değerlendirme Formu` · `Soru Seti` · `Açılış ve Kapanış` (sorulmaması gerekenler, son kontrol) · `.md` soru seti

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek medeni hâl sorusu sorar, bir yetkinliği atlar ve
süreyi aşar. Kodun bunları yakalayıp düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve mülakat sürecinin hukuka uygunluğu kullanıcının
sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
