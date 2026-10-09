# Standart-Fiili Maliyet Sapma Analizi · Kod Bloğu

> Muhasebe › Maliyet Muhasebesi Uzmanı · Workers / Workless

Standart maliyet ile fiili maliyet arasındaki farkı klasik sapma bileşenlerine ayırır. Sapmaları Tek Düzen Hesap
Planı (7/A) fark hesaplarıyla raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

Pozitif tutar **aleyhte** (fiili > standart), negatif tutar **lehtedir**. Standart miktar ve saat, dönemin fiili
üretimi için hesaplanır: Σ ürün (üretim × birim standart).

| Maliyet | Sapma | Formül | Hesap |
|---|---|---|---|
| Direkt ilk madde ve malzeme | Fiyat | (fiili fiyat − standart fiyat) × fiili miktar | 711 |
| Direkt ilk madde ve malzeme | Miktar | (fiili miktar − standart miktar) × standart fiyat | 712 |
| Direkt işçilik | Ücret | (fiili ücret − standart ücret) × fiili saat | 721 |
| Direkt işçilik | Süre | (fiili saat − standart saat) × standart ücret | 722 |
| Genel üretim giderleri | Bütçe | fiili GÜG − esnek bütçe (fiili saat) | 731 |
| Genel üretim giderleri | Verimlilik | (fiili saat − standart saat) × standart GÜG oranı | 732 |
| Genel üretim giderleri | Kapasite | esnek bütçe (fiili saat) − fiili saat × standart GÜG oranı | 733 |

GÜG "üç sapma yöntemi" ile iş merkezi bazında hesaplanır:
- Standart GÜG oranı = değişken oran + bütçelenen sabit GÜG ÷ normal kapasite.
- Esnek bütçe = bütçelenen sabit GÜG + fiili saat × değişken oran.
- Yük tabanı, iş merkezinin direkt işçilik saatidir. İşçilik kalemi adı ile GÜG iş merkezi adı aynı olmalıdır.

Bileşenlerin toplamı fiili − standart maliyete eşittir; rapor bu mutabakatı gösterir.

| Kontrol | Önem |
|---|---|
| Üretilen ürünün standart kartı yok | Yüksek |
| Standartta olan kalemin fiili kaydı yok | Yüksek |
| Kalem sapması standardın `--esik` %'sini aşıyor (3 katı Yüksek) | Orta / Yüksek |
| Standart kartta olmayan fiili kalem; tanımsız iş merkezi | Orta |
| Fiili saat normal kapasitenin %80'inin altında | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 2 ürün, 4 malzeme, 2 iş merkezi
python main.py --standartlar s.xlsx --uretim u.xlsx --fiili f.xlsx --gug gug.xlsx --esik 5
```

| Dosya | Sütunlar |
|---|---|
| Standartlar | Ürün, Tür (Malzeme / İşçilik), Kalem, Birim, Standart Miktar (birim ürün başına), Standart Fiyat (TL / birim veya saat) |
| Üretim | Ürün, Üretim Miktarı |
| Fiili | Tür (Malzeme / İşçilik / GÜG), Kalem, Fiili Miktar (GÜG'de boş), Fiili Tutar |
| GÜG (isteğe bağlı) | İş Merkezi, Değişken Oran (TL/saat), Bütçelenen Sabit GÜG, Normal Kapasite (saat) |

## Çıktı

`maliyet_sapma_analizi.xlsx`:
- `Sapma Özeti`: hesap kodu bazında sapmalar, toplam, standart ve fiili maliyet, mutabakat satırı, grafik.
- `Malzeme` ve `İşçilik`: kalem bazında standart / fiili miktar, fiyat ve tutar; iki sapma ve sapma %'si.
- `GÜG`: iş merkezi bazında oranlar, saatler, esnek bütçe ve üç sapma.
- `Standart Maliyet Kartı`: ürün başına malzeme, işçilik ve GÜG standart maliyeti.
- `Uyarılar`: boş "Açıklama / Aksiyon" sütunu.

## Dikkat

- **Fiyat sapması kullanılan miktar üzerinden hesaplanır.** Satın alma anında ayrıştırma (satın alınan miktar
  üzerinden) yapıyorsanız sonuç farklı olur.
- **Yarı mamul ve stok değişimi:** Dönem başı / sonu yarı mamul stokları için eşdeğer üretim miktarını üretim dosyasına
  siz girmelisiniz.
- **Sapmaların kapatılması:** Sapmaların satılan mamul maliyeti ile stoklar arasında dağıtımı muhasebe politikasına
  bağlıdır; paket kayıt yapmaz.
- **Örnek veri:** Örnek standartlar ve fiili değerler kurgusaldır.

## Testler

Şunlar test edilir:
- Malzeme ve işçilik sapmaları (elle hesaplanmış değerlerle); GÜG üç sapma (lehte ve aleyhte).
- Toplam sapma = fiili − standart mutabakatı.
- Standart dışı kalem, eksik kart, eksik fiili kayıt ve üretimi olmayan ürün; Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
