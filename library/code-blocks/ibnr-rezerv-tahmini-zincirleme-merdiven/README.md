# IBNR Rezerv Tahmini (Zincirleme Merdiven) · Kod Bloğu

> Sigortacılık › Aktüerya › Aktüerya Uzmanı · Workers / Workless

Hasar gelişim üçgeninden **zincirleme merdiven** (chain ladder) yöntemiyle nihai hasarı ve gerçekleşmiş ancak
rapor edilmemiş (IBNR) / henüz ödenmemiş hasar karşılığını hesaplar. **Mack (1993)** dağılımdan bağımsız
standart hatasını kaza yılı bazında ve toplamda verir. İnternete bağlanmaz.

## Ne yapar?

**Üçgenin kurulması**
- **Hazır üçgen:** Kümülatif veya artımlı (`--artimli`) matris olarak verilebilir.
- **Ödeme veya ihbar dökümü:** Kaza tarihi, işlem tarihi ve tutar sütunlarından yıllık ya da çeyreklik
  (`--donem ceyrek`) üçgen kurulur. `--degerleme` ile değerleme tarihinden sonraki işlemler dışarıda bırakılır.

**Gelişim faktörleri**
- **Ortalama türü:** Hacim ağırlıklı (varsayılan) veya basit ortalama.
- **Son N dönem:** `--son-n` ile yalnız son N kaza dönemi kullanılır.
- **Elle seçim:** `--faktor 8=1,07` ile tek tek faktör girilir.
- **Kuyruk faktörü:** `--kuyruk` ile eklenir.

**Mack standart hatası**
- **Süreç ve tahmin hatası:** Kaza dönemi bazında hesaplanır ve değişim katsayısı verilir.
- **Toplam hata:** Kaza dönemleri arası korelasyonu (ortak faktörler) içerir.
- **Son dönemin σ²'si:** Mack'in dışdeğerlemesiyle hesaplanır.

**Tanı tabloları**
- **Bağlantı oranları:** Kullanılan faktörden 2 standart sapmadan fazla uzak oranlar kırmızı işaretlenir. Bu
  oranlar büyük hasar veya süreç değişikliği belirtisi olabilir.
- **Gelişmişlik yüzdesi:** Her kaza dönemi için 1/CDF olarak verilir.

## Doğrulama

Örnek veri, aktüerya literatüründe standart olan **Taylor & Ashe (1983)** üçgenidir. Bu üçgen Mack (1993) ve R
ChainLadder paketinde `GenIns` adıyla kullanılır. Paket şu sonuçları birebir üretir:

| | Değer |
|---|---|
| Toplam IBNR | 18.680.856 |
| Toplam Mack standart hatası | 2.447.095 (%13,1) |
| Son kaza yılı IBNR / std. hata | 4.625.811 / 1.363.155 |

Testler bu değerleri kontrol eder. Kaza yılları gösterim için 2016–2025 olarak etiketlenmiştir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                  # Taylor-Ashe örneği
python main.py --ucgen ucgen.xlsx                               # kümülatif üçgen
python main.py --ucgen ucgen.xlsx --artimli --son-n 5 --kuyruk 1,02
python main.py --hasarlar odemeler.xlsx --donem ceyrek --degerleme 30.09.2026
python main.py --ucgen ucgen.xlsx --faktor 8=1,07 9=1,015 --tur "gerçekleşen hasar"
```

## Çıktı

`Özet` (kaza dönemi bazında nihai hasar, IBNR, std. hata, grafik) · `Faktörler` · `Üçgen` (tahmin edilen hücreler
renkli) · `Bağlantı Oranları`

## Dikkat

- **Varsayımlar:** Yöntem, geçmiş gelişim örüntüsünün geleceğe taşınacağını varsayar. Büyük ve katastrofik
  hasarlar, hasar enflasyonundaki değişimler, ödeme ve rezerv uygulamasındaki değişiklikler sonucu bozar. Bu
  durumlarda veri ayıklanmalı veya ayrı analiz yapılmalıdır.
- **Mevzuat:** Türkiye'de teknik karşılıklara ilişkin yönetmelik ve SEDDK düzenlemeleri, IBNR hesabında veri
  seçimi, büyük hasar ayıklama ve faktör seçimi gibi ek kurallar içerir. Bu araç yöntemi uygular; yasal karşılık
  hesabı sorumlu aktüer tarafından değerlendirilip onaylanmalıdır.
- **Kuyruk belirsizliği:** Kuyruk faktörü kullanıldığında Mack standart hatası kuyruk belirsizliğini içermez.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
