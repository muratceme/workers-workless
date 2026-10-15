# Vitrin Yönergesi Hazırlama · AI Agent

> Görsel Düzenleme › Görsel Düzenleme Uzmanı · Workers / Workless

Kampanya ve ürün listesinden **mağazalara gönderilecek adım adım vitrin düzenleme yönergesi** hazırlar. Her vitrin tipi
(büyük, küçük, köşe …) için ayrı yönerge yazılır. Manken ve podyum kapasitesi, zorunlu ürünler ve kombinler kodla
kontrol edilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Üç dosya okunur:
   - Kampanya bilgisi
   - Ürün listesi (SKU, grup, renk, Zorunlu / Opsiyonel)
   - Vitrin tipleri (manken ve podyum sayısı, mağazalar)
2. **Yönerge (model):** Her vitrin tipi için sıralı adımlar, manken kombinleri, podyum yerleşimi, afiş metni ve kontrol
   listesi yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o vitrin tipinin yönergesi modele geri gönderilir (varsayılan en
   fazla 2 tur):
   - Listede olmayan SKU kullanılmış
   - Manken ya da podyum sayısı vitrinin kapasitesini aşıyor ya da numara tekrarlanıyor
   - "Zorunlu" bir ürün o vitrinde kullanılmamış
   - Bir mankende farklı gruplardan ürün var (ör. kadın kaban + erkek pantolon)
   - Vitrindeki ürünlerin fiyat etiketiyle ilgili adım yok
   - Afişte ya da adımlarda kampanya bilgisinde olmayan sayı / tarih var (ör. uydurulmuş "%40 indirim")
4. **Dağıtım (kod):** Kontrolden geçen her vitrin tipi için ayrı bir .txt yönerge yazılır. Excel'de hangi mağazaya hangi
   dosyanın gideceği listelenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                         # örnek sonbahar vitrini
python agent.py --kampanya kampanya.txt --urunler urunler.xlsx --vitrinler vitrin_tipleri.xlsx
```

## Çıktı

- `cikti/vitrin_yonergesi_buyuk.txt`, `cikti/vitrin_yonergesi_kucuk.txt` …: mağazalara gönderilecek yönergeler.
- `cikti/vitrin_yonergesi.xlsx`:
  - `Dağıtım`: mağaza → vitrin tipi → yönerge dosyası, kontrol sonucu.
  - Her vitrin tipi için ayrıntı sayfası.

## Dikkat

- **Onay ve stok.** Yönergeler görsel düzenleme sorumlusu onaylamadan gönderilmemelidir. Mağazanın stoğunda olmayan ürün
  için yedek ürün belirleyin.
- **Fiyat etiketi.** Vitrinde sergilenen ürünlerin fiyatlarının görünür olması gerekir; kod yönergede fiyat etiketi adımı
  olduğunu kontrol eder, uygulamayı mağaza ekibi yapar.
- **Örnek veri:** Kampanya, ürünler ve mağazalar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk yönergelerde bilerek:
- büyük vitrine 4 yerine 5 manken koyar ve fiyat etiketi adımını atlar,
- küçük vitrinde kadın kabanla erkek pantolonu aynı mankene giydirir,
- zorunlu erkek montu unutur,
- afişe "%40 indirim" ekler.

Kodun bunları yakaladığı, mağaza adlarının modele gönderilmediği ve kontrolden geçen yönergelerin vitrin tipine göre
dosyalandığı doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve mağaza uygulamaları kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
