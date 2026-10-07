# Kıdem ve İhbar Tazminatı Hesaplama · Kod Bloğu

> İnsan Kaynakları › Bordro ve Özlük İşleri Uzmanı · Workers / Workless

İşe giriş-çıkış tarihi ve ücretten **kıdem tazminatını** ve **ihbar tazminatını** hesaplar; vergi
kesintileriyle net tutarları çıkarır. Tek kişi veya toplu çıkış listesi için çalışır. İnternete bağlanmaz.

## Kurallar

| Kalem | Kural |
|---|---|
| Hizmet süresi | Giriş ve çıkış günleri dahil yıl / ay / gün |
| Giydirilmiş ücret | Çıplak brüt + düzenli aylık yan ödemeler (yemek, yol vb.) + yıllık ikramiye ÷ 12 |
| Kıdem | Her tam yıl için 30 günlük giydirilmiş ücret; artan ay ÷ 12, gün ÷ 365 oranında (1475 s. Kanun md. 14). En az 1 yıl hizmet gerekir |
| Kıdem tavanı | Fesih tarihindeki tavan: 01.01-30.06.2026 → 64.948,77 TL · 01.07-31.12.2026 → 73.729,87 TL (2025 dönemleri de tanımlı) |
| Kıdem kesintisi | Gelir vergisinden istisna (GVK md. 25/7); yalnız damga vergisi (binde 7,59) |
| İhbar süresi | 6 aydan az 2 hafta · 6 ay-1,5 yıl 4 hafta · 1,5-3 yıl 6 hafta · 3 yıldan fazla 8 hafta (4857 s. Kanun md. 17) |
| İhbar tutarı | Giydirilmiş günlük ücret (aylık ÷ 30) × gün; tavan uygulanmaz |
| İhbar kesintisi | Gelir vergisi (çıkış yılındaki kümülatif matraha göre dilim) + damga vergisi |

Parametreler `tr_parametreler.json` dosyasındadır ve her biri kaynaklıdır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py --giris 15.03.2019 --cikis 30.09.2026 --brut 60000 --yan-odeme 4500 --ikramiye 60000 --kumulatif 450000
python main.py --girdi cikislar.xlsx
python main.py                     # örnek listeyle dener
```

Toplu girdi sütunları: **Ad Soyad, İşe Giriş, İşten Çıkış, Brüt** (zorunlu); Yan Ödeme, Yıllık İkramiye,
Kümülatif Matrah, Kıdem Hakkı (E/H), İhbar Hakkı (E/H) (isteğe bağlı). Tarihler `GG.AA.YYYY` veya `YYYY-AA-GG`.

## Önemli

Kıdem ve ihbar hakkının doğup doğmadığı **fesih nedenine** bağlıdır (raporun "Kurallar" sayfasında özetlenmiştir).
Bu araç hak durumunu belirlemez; "Kıdem Hakkı / İhbar Hakkı" sütunlarıyla siz belirtirsiniz. Kullanılmayan
yıllık izin ücreti dahil değildir. Sonuçlar bilgilendirme amaçlıdır; uyuşmazlıkta hukuki destek alın.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
