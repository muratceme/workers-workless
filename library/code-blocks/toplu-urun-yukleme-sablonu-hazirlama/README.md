# Toplu Ürün Yükleme Şablonu Hazırlama · Kod Bloğu

> E-ticaret › Pazaryeri Yönetimi › Pazaryeri Uzmanı · Workers / Workless

Ürün ana veri dosyanızı pazaryerinin **toplu ürün yükleme şablonuna** dönüştürür ve yüklemeden önce
kontrol eder. Eksik zorunlu alanları, hatalı barkodları, eşleşmeyen kategori ve renkleri, varyant hatalarını
listeler. İnternete bağlanmaz; pazaryerine hiçbir şey yüklemez.

## Neden eşleştirme dosyası?

Her pazaryerinin şablonu farklıdır ve zamanla değişir. Bu paket belirli bir pazaryerinin şablonunu taklit
etmez. Satıcı panelinden indirdiğiniz şablonu, sütunları bir **eşleştirme dosyasında** (JSON) tanımlayarak
doldurur. Bir kez hazırlanan eşleştirme her yüklemede tekrar kullanılır.

Örnek dosya: `ornek_veri/eslestirme.json` (kurgusal "Örnek Pazaryeri" için). Her şablon sütunu için şunlar
tanımlanabilir:

| Anahtar | Anlamı | Örnek |
|---|---|---|
| `kaynak` | Ana verideki sütun | `"kaynak": "Liste Fiyatı"` |
| `sabit` | Her üründe aynı değer | `"sabit": "TRY"` |
| `sablon` | Birleşik metin; `{Sütun}` yer tutucuları | `"{Marka} {Ürün Adı} {Renk}"` |
| `bol` + `sira` | Tek hücredeki listeden n'inci değer | Görseller `\|` ile ayrılmış, `"sira": 2` |
| `eslestirme` + `sutun` | Değeri eşleştirme tablosundan çevir | Bizim kategori → pazaryeri kategori adı / ID |
| `zorunlu` | Boş olamaz | |
| `tur` | `para`, `sayi`, `tamsayi` | |
| `izinli` | İzin verilen değerler | KDV: `[0, 1, 10, 20]` |
| `en_az`, `en_fazla` | Karakter sınırı | Ürün adı en fazla 100 |
| `kontrol` | `gtin` (barkod kontrol hanesi) veya `url` (https görsel adresi) | |
| `benzersiz` | Tüm dosyada tekrar edemez | Barkod, stok kodu |

**Kurallar:**
- **Fiyat:** Satış fiyatı liste/piyasa fiyatından yüksek olamaz.
- **Varyant:** Aynı model kodundaki ürünlerde marka, kategori ve açıklama aynı olmalı; renk-beden ikilisi tekrar
  etmemeli.

**Eşleştirmeler:**
- Kategori gibi uzun listeler için CSV dosyası kullanılır (`kategori_eslestirme.csv`).
- Renk gibi kısa listeler doğrudan JSON'a yazılabilir.
- Karşılığı olmayan değerler raporun `Eşleşmeyen Değerler` sayfasına çıkar. Bunları doldurup eşleştirme
  tablosuna ekleyin.

## Kullanım

```bash
pip install -r requirements.txt
python main.py                                                      # örnek: 13 varyant, kurgusal şablon
python main.py --girdi urunler.xlsx --sablon pazaryeri_sablonu.xlsx --eslestirme eslestirme.json
python main.py --girdi urunler.xlsx --sablon sablon.xlsx --eslestirme eslestirme.json --hepsi   # hatalılar da yazılsın
```

- **Şablon `.xlsx` ise:** Dosya kopyalanır. İlk sayfa, başlık satırının (`sablon_baslik_satiri`) altından
  doldurulur.
- **Açıklama satırları:** Şablonda başlığın altında açıklama satırları varsa JSON'a
  `"veri_baslangic_satiri": 4` gibi bir değer ekleyin. Talimat sayfaları olduğu gibi kalır.
- **Barkod:** Barkod, model kodu ve stok kodu metin olarak yazılır; baştaki sıfırlar kaybolmaz.

## Çıktı

- **Yükleme dosyası:** `yukleme_<pazaryeri>.xlsx`, yalnız hatasız ürünleri içerir.
- **Kontrol raporu:** `yukleme_kontrol_raporu.xlsx`:
  - `Özet`: sayılar ve en sık sorunlar.
  - `Hatalar`: kaynak satır numarasıyla, "Düzeltildi mi?" sütunu.
  - `Eşleşmeyen Değerler`.
  - `Ayar Kontrolü`: şablonda olup eşleştirmede olmayan sütunlar ve ana veride bulunamayan sütunlar.

## Dikkat

- **Güncel kurallar:** Zorunlu alanlar, karakter sınırları, kategoriye özel özellikler ve görsel kuralları
  pazaryerine ve kategoriye göre değişir. Örnek eşleştirmedeki değerler örnektir; satıcı panelindeki güncel
  şablonu ve kılavuzu esas alın.
- **KDV oranı:** Ürününüze uygulanacak KDV oranını mali müşavirinizle teyit edin. Örnekteki `0, 1, 10, 20`
  listesi yalnız biçim kontrolüdür.
- **Yükleme sonrası:** Yüklemeden sonra pazaryerinin kendi hata raporunu da kontrol edin.
- **Örnek veri:** Örnek ürünler, barkodlar ve görsel adresleri kurgusaldır.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçlar kontrol edilmeden yükleme yapılmamalıdır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
