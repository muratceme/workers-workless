# Puantaj Kontrolü · Kod Bloğu

> İnsan Kaynakları › Bordro ve Özlük İşleri Uzmanı · Workers / Workless

Aylık puantaj çizelgesini bordro öncesi kontrol eder. Hatalı satırları, İş Kanunu sınırlarını aşan çalışmaları
ve bordroda karar gerektiren durumları listeler. SGK prim gününü ve eksik gün nedenini önerir. İnternete bağlanmaz.

## Puantaj kodları

| Kod | Anlamı | Ücretli | SGK eksik gün nedeni |
|---|---|---|---|
| `Ç` (X, 1) | Çalıştı; fazla mesai `Ç+2`, `X 2,5` | Evet | |
| `HT` | Hafta tatili | Evet | |
| `GT` | Genel tatil (çalışılmadı) | Evet | |
| `Yİ` | Yıllık ücretli izin | Evet | |
| `Mİ` | Mazeret / yasal ücretli izin | Evet | |
| `R` | Rapor (istirahat) | Hayır | 01 |
| `Üİ` | Ücretsiz izin | Hayır | 21 |
| `D` | Devamsızlık | Hayır | 15 |

## Kontroller

| Seviye | Kontrol |
|---|---|
| **Hata** | Tanımsız kod · boş gün · işe giriş öncesi/çıkış sonrası doldurulmuş gün · ayda olmayan gün · genel tatil olmayan güne GT · toplam sütunlarının (çalışılan gün, fazla mesai, SGK gün) günlerle tutmaması · 30'u aşan prim günü |
| **Yasal** | 7 gün üst üste çalışma (md. 46 hafta tatili) · günlük 11 saati aşan çalışma (md. 63) · yıllık 270 saati aşan fazla çalışma (md. 41; "Yıl Başından FM" sütunu verilirse) |
| **Dikkat** | Yıllık izne denk gelen genel tatil (md. 56: izinden sayılmaz) · devamsızlık olan haftada ücretli hafta tatili (md. 46 koşulu) · yıllık izin bakiyesinin aşılması · 31/28 çeken aylarda 1 günlük SGK gün farkı |
| **Bilgi** | Genel tatilde çalışma (md. 47: ayrıca bir günlük ücret) |

**SGK prim günü**
- **Tam ay:** 30 − eksik gün.
- **Ay içinde giriş veya çıkış:** İstihdam edilen gün sayısı (en çok 30) − eksik gün.
- **Eksik gün nedeni:** Tek bir neden varsa onun kodu (01, 15 veya 21) önerilir. Birden fazla neden varsa 12
  (Birden fazla nedenle) önerilir.

**Genel tatiller**
- **Hesaplanan tarihler:** Sabit tatiller ile 2025–2027 Ramazan ve Kurban Bayramları dahildir.
- **Arifeler:** 13.00'ten itibaren yarım gün tatil sayılır, bu yüzden GT kodu verilmez.
- **Ek tatiller:** İdari izin ve başka yılların bayram tarihleri `--tatil` dosyasıyla eklenir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                           # örnek Mayıs 2026 puantajıyla
python main.py --girdi puantaj.xlsx --donem 2026-09
python main.py --girdi puantaj.xlsx --gunluk-saat 9 --tatil ek_tatiller.csv
```

- **Dönem:** Dosyanın başlığında "Mayıs 2026" gibi bir ifade varsa `--donem` gerekmez.
- **Günlük normal saat:** Varsayılan 7,5 saattir (45 saat / 6 gün). 5 gün çalışılan işyerlerinde `--gunluk-saat 9`
  verin.
- **Toplam sütunları:** Çalışılan Gün, Fazla Mesai Saati, SGK Gün, Kalan Yıllık İzin ve Yıl Başından FM isteğe
  bağlıdır. Verilenler günlük kodlardan hesaplanan değerlerle karşılaştırılır.

## Çıktı

`Özet` (ayın tatilleri dahil) · `Bulgular` · `Personel Özeti` · `Puantaj` (renkli görünüm) · `Kodlar`

## Dikkat

- **Hafta tatili:** 7 gün kuralı yalnızca ay içinde görülebilen günler için uygulanır. Ay geçişlerini önceki ayın
  puantajıyla birlikte kontrol edin.
- **Devamsızlık haftası:** Devamsızlık olan haftanın hafta tatili ücreti, toplu sözleşme ve işyeri uygulamasına
  göre değişebilir.
- **Bildirim öncesi:** Bordro ve SGK bildirimi öncesinde uzman kontrolü yapın. Eksik gün belgelerini (rapor,
  tutanak, izin dilekçesi) saklayın.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
