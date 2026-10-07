# Mali Tablo Rasyo Analizi · Kod Bloğu

> Finans › Finansal Analist · Workers / Workless

Muhasebe programından (Logo, Mikro, Netsis, Luca vb.) alınan **mizanlardan** Tekdüzen Hesap Planı'na göre özet
bilanço ve gelir tablosu oluşturur. Likidite, finansal yapı, faaliyet ve kârlılık rasyolarını dönemler arası
karşılaştırmalı hesaplar ve her rasyonun iyileştiğini ya da kötüleştiğini işaretler. İnternete bağlanmaz.

## Rasyolar

| Grup | Rasyolar |
|---|---|
| Likidite | Cari oran, asit-test oranı, nakit oranı |
| Finansal yapı | Kaldıraç, özkaynak/aktif, borç/özkaynak, KVYK/toplam yabancı kaynak, duran varlık/özkaynak |
| Faaliyet | Alacak, stok ve ticari borç devir hızları ve süreleri, **nakit dönüşüm süresi**, aktif devir hızı |
| Kârlılık | Brüt, esas faaliyet ve net kâr marjı, ROA, ROE, faiz karşılama oranı |

- **Ortalama bakiye:** Önceki dönem verildiyse devir hızlarında ve ROA/ROE'de ortalama bakiye
  ((önceki + son) / 2) kullanılır.
- **Ara dönem:** 9 aylık mizan gibi ara dönemlerde `--gun` ile dönem gün sayısını verin; devir süreleri buna
  göre hesaplanır.
- **Referans değerler:** Raporda yer alan genel değerlerdir (ör. cari oran 1,5 – 2). Sektöre göre değişir.

## Mizan nasıl okunur?

- Üstte firma ve dönem başlığı olabilir; başlık satırı "Hesap Kodu" sütunundan bulunur.
- Bakiye `Borç Bakiye` / `Alacak Bakiye`, `Borç` / `Alacak` ya da işaretli `Bakiye` sütunlarından alınır.
- **Ana hesap** (3 hane) satırı varsa o kullanılır. Yoksa en alt kırılımlar toplanır. Ara seviyeler (ör.
  `320.01`) çift sayılmaz.
- Mizan **kapanış kayıtlarından önce** alınmalıdır, yani 6'lı hesaplar açık olmalıdır. Kapanmış mizanda gelir
  tablosu boş kalır ve dönem kârı 590/591'den alınır.
- Mizan denk değilse, bilanço denk değilse ya da 7'li hesaplarda yansıtma yapılmamış bakiye kaldıysa uyarı
  verilir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                   # örnek 2025 ve 2026 mizanlarıyla
python main.py --mizan 2025=mizan_2025.xlsx 2026=mizan_2026.xlsx
python main.py --mizan 2025=mizan_2025.xlsx 2026/09=mizan_eylul.xlsx --gun 365 273
```

## Çıktı

`Rasyolar` · `Bilanço` · `Gelir Tablosu` (grafikli) · `Dikey Analiz`

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
