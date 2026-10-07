# Risk Değerlendirmesi (Fine-Kinney) · Kod Bloğu

> İş Sağlığı ve Güvenliği › İş Güvenliği Uzmanı · Workers / Workless

Tehlike listesinden Fine-Kinney risk skorunu (**R = Olasılık × Frekans × Şiddet**) ve risk sınıfını
hesaplar; sonucu sınıfa göre renklendirilmiş, en yüksek riskten başlayarak sıralanmış bir Excel
risk değerlendirme tablosuna döker. Önlem sonrası (kalıntı) riski de karşılaştırır. İnternete bağlanmaz.

## Skalalar

| Olasılık (O) | | Frekans (F) | | Şiddet (Ş) | |
|---|---|---|---|---|---|
| 10 | Kuvvetle beklenir | 10 | Sürekli (saatlik) | 100 | Facia (birden fazla ölüm) |
| 6 | Oldukça mümkün | 6 | Sıklıkla (günlük) | 40 | Felaket (ölümlü kaza) |
| 3 | Olağan dışı fakat olabilir | 3 | Ara sıra (haftalık) | 15 | Çok ciddi (iş günü kaybı) |
| 1 | Çok uzak ihtimal | 2 | Nadir (aylık) | 7 | Ciddi (dış ilk yardım) |
| 0,5 | İhtimal dahilinde fakat beklenmez | 1 | Seyrek (yıllık) | 3 | Önemli (dahili ilk yardım) |
| 0,2 | Pratik olarak imkânsız | 0,5 | Oldukça seyrek | 1 | Fark edilebilir |
| 0,1 | Neredeyse imkânsız | | | | |

| Risk skoru | Sınıf | Eylem |
|---|---|---|
| R < 20 | Kabul edilebilir | Acil eylem gerekmeyebilir |
| 20 ≤ R < 70 | Olası risk | Eylem planına alınmalı, gözetim altında tutulmalı |
| 70 ≤ R < 200 | Önemli risk | Dikkatle izlenmeli, eylem planına alınmalı |
| 200 ≤ R < 400 | Yüksek risk | Kısa vadeli eylem planına alınmalı |
| R ≥ 400 | Çok yüksek risk | Tolerans gösterilemez; faaliyet durdurulmalı |

Kaynak: Kinney ve Wiruth (1976); Türkçe akademik uygulamalardaki skala ve sınıflar.
Kurumunuz farklı bir eşik seti kullanıyorsa `main.py` içindeki `SINIFLAR` listesini düzenleyin.

## Kontroller

- Skalada olmayan puan (ör. şiddet 5) → hesaplanmaz, uyarı verilir
- Önemli ve üzeri riskte önerilen önlem veya sorumlu boş → uyarı
- Termini geçmiş aksiyon → uyarı
- Önlem sonrası risk önlem öncesinden yüksekse → uyarı

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                   # örnek tehlike listesiyle dener
python main.py --girdi tehlike_listesi.xlsx
```

Sütunlar: **Tehlike, Olasılık, Frekans, Şiddet** (zorunlu); Bölüm, Risk, Mevcut Önlemler, Önerilen Önlemler,
Sorumlu, Termin, Önlem Sonrası Olasılık, Önlem Sonrası Frekans, Önlem Sonrası Şiddet (isteğe bağlı).

> Risk değerlendirmesi 6331 sayılı Kanun kapsamında risk değerlendirme ekibince yapılmalıdır; bu araç
> hesaplama ve raporlamayı kolaylaştırır, uzman değerlendirmesinin yerini tutmaz.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
