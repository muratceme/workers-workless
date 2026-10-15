# Trend Raporu Özeti · AI Agent

> Tasarım › Moda Tasarımcısı · Workers / Workless

Sezon trend raporlarını koleksiyonunuza **uygulanabilir renk, kumaş, silüet, detay ve desen maddelerine** indirir. Her
madde rapordan aynen alıntıyla gelir. Koleksiyon brifinize uymayan trendler gerekçesiyle ayrı listelenir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Trend raporu satır (paragraf) satır numaralanır. Koleksiyon brifinden ürün grupları ve
   kaçınılacaklar okunur. Raporda hangi kategorilerin geçtiği anahtar kelimelerle belirlenir.
2. **Özet (model):** Her trend için kategori, kısa ad, rapordan alıntı, koleksiyona uygulama, ürün grupları ve öncelik
   yazılır. Brife uymayan trendler "uygulanmayacaklar" bölümüne gider.
3. **Denetim (kod):** Aşağıdakilerden biri varsa özet modele geri gönderilir (varsayılan en fazla 2 tur):
   - Alıntı dayandığı rapor satırında aynen geçmiyor
   - Ürün grubu brifte yok
   - Pantone (TCX / TPX) kodu ya da sayı raporda / brifte yok (renk kodu uydurulamaz)
   - Raporda geçen bir kategoride hiç madde yok
   - Brifteki kaçınılacaklardan biri (ör. payet, deri) uygulama maddesinde önerilmiş

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                  # örnek (kurgusal) trend notları + brif
python agent.py --rapor trend_raporu.pdf --brif koleksiyon_brifi.txt
```

Brif örneği:

```
Sezon: İlkbahar/Yaz 2027
Hedef Müşteri: 25-40 yaş arası, şehirde çalışan kadın
Ürün Grupları: Gömlek, Pantolon, Elbise, Etek, Tişört
Fiyat Segmenti: Orta
Ana Kumaşlar: Pamuk, keten, viskon
Kaçınılacaklar: Payet, metalik kumaş, deri
```

## Çıktı

- `cikti/trend_ozeti.xlsx`:
  - `Maddeler`: kategori, trend, rapordan alıntı, satır, koleksiyona uygulama, ürün grupları, öncelik ve "Karar" sütunu.
  - `Uygulanmayacaklar`: gerekçesiyle.
- `cikti/trend_ozeti.txt`: kategori bazında özet.

## Dikkat

- **Lisanslı raporlar.** Ücretli trend servislerinin raporlarını bir yapay zekâ sağlayıcısına göndermeden önce
  abonelik sözleşmenizin buna izin verdiğini kontrol edin. Agent rapor metnini kopyalayıp yayımlamaz; yalnız kısa
  alıntılar kullanır.
- **Renk doğrulaması.** Pantone kodlarını fiziksel renk kartlarıyla, kumaşları tedarikçi numuneleriyle doğrulayın.
- **Taranmış PDF.** Görüntü olarak taranmış raporlardan metin çıkarılamaz; önce OCR uygulayın.
- **Örnek veri:** Trend notları ve brif kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk özette bilerek:
- alıntıyı rapordaki gibi yazmaz,
- raporda olmayan bir Pantone kodu (17-1463 TCX) uydurur,
- brifte olmayan "Ceket" grubunu ekler,
- silüet kategorisini atlar,
- kaçınılacaklar listesindeki payet ve metalik kumaşı önerir.

Kodun bunları yakalayıp düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve tasarım kararları kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
