# Doluluk, ADR ve RevPAR Raporu · Kod Bloğu

> Otel › Ön Büro / Gelir Yönetimi · Workers / Workless

Oda satışlarından otelin temel performans göstergelerini **günlük, aylık, haftanın günü ve kanal** bazında
hesaplar ve **geçen yılla** karşılaştırır. İnternete bağlanmaz.

| Gösterge | Formül |
|---|---|
| Doluluk | Ücretli satılan oda ÷ kullanılabilir oda |
| Toplam doluluk | (Ücretli + ücretsiz + kendi kullanım) ÷ kullanılabilir oda |
| ADR (ortalama oda fiyatı) | Oda geliri ÷ ücretli satılan oda |
| RevPAR | Oda geliri ÷ kullanılabilir oda = Doluluk × ADR |

- **Ücretsiz ve kendi kullanım odaları:** Ücretsiz (complimentary) ve kendi kullanım (house use) odaları ADR'yi
  düşürmesin diye satılan odaya katılmaz. Toplam dolulukta ayrıca gösterilir.
- **Kullanım dışı odalar:** Arızalı odalar (OOO) varsayılan olarak kullanılabilir odadan düşülmez;
  `--ooo-dus` ile düşülür.
- **Geçen yıl karşılaştırması:** Haftanın aynı gününe hizalanır (364 gün önce). Böylece cumartesi cumartesiyle
  karşılaştırılır.
- **Rezervasyon listesi:** Konaklamalar gecelere açılır (giriş dahil, çıkış hariç). İptal ve no-show
  rezervasyonlar çıkarılır. `Ücret Tipi` "Ücretsiz" olanlar ücretsiz oda sayılır. Kanal sütunu varsa kanal ve
  segment kırılımı çıkarılır.
- **Overbooking:** Dolu oda sayısı oda kapasitesini aşan günler uyarılır.

> **Oda geliri** KDV ve konaklama vergisi hariç olmalıdır. Paket fiyatlarda (kahvaltı, yarım pansiyon vb.)
> yalnız oda payı kullanılmalıdır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                             # örnek: 40 odalı otel, Eylül 2026
python main.py --girdi gunluk_istatistik.xlsx --oda 120
python main.py --girdi rezervasyonlar.xlsx --oda 120 --baslangic 01.09.2026 --bitis 30.09.2026 --ooo-dus
```

## Çıktı

`Aylık` · `Günlük` (doluluk, ADR ve RevPAR grafikleri) · `Haftanın Günü` · `Kanal` · `Bilgi`

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
