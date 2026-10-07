# Kumaş Kontrol (4 Puan Sistemi) Raporu · Kod Bloğu

> Tekstil › Kalite Kontrol · Workers / Workless

Kumaş kontrol makinesinde tutulan **top bazındaki hata kayıtlarından** 4 puan sistemine göre puan hesaplar.
Puanı 100 yard² (ve 100 m²) başına çevirir; topu ve partiyi kabul veya ret olarak raporlar. İnternete bağlanmaz.

## Puanlama (ASTM D5430'da tanımlanan yaygın uygulama)

| Hata uzunluğu | Puan |
|---|---|
| ≤ 3 inç (7,62 cm) | 1 |
| 3 – 6 inç (15,24 cm) | 2 |
| 6 – 9 inç (22,86 cm) | 3 |
| > 9 inç | 4 |
| Delik / açıklık ≤ 1 inç (2,54 cm) | 2 |
| Delik / açıklık > 1 inç | 4 |

- **Yard sınırı:** Bir yarda (0,9144 m) **en fazla 4 puan** verilir.
- **Sürekli hata:** 1 yarddan uzun boyuna hata, uzandığı **her yard için 4 puan** alır.
- **Formül:** `Puan / 100 yd² = toplam puan × 36 × 100 ÷ (kesilebilir en [inç] × uzunluk [yd])`
- **Kabul:** Varsayılan top kabul sınırı **40 puan / 100 yd²**'dir. Alıcı kriterlerine göre `--top-esik` ve
  `--parti-esik` ile değiştirilir.
- **Parti puanı:** Alan ağırlıklıdır (toplam puan ÷ toplam alan).
- **Ek kontroller:**
  - Ölçülen uzunluğun etiketten kısa olması
  - Kesilebilir enin sipariş eninden dar olması (tolerans `--en-tolerans`)
  - Art arda 3 veya daha fazla yardda 4 puan (kesim planında dikkat)

Formül, yayımlanmış örnekle test edilmiştir: 120 yd × 45 inç topta 22 puan = 14,67 puan / 100 yd².

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                          # örnek 5 top, 2 parti (metrik)
python main.py --toplar toplar.xlsx --hatalar hata_kayitlari.xlsx
python main.py --toplar toplar.xlsx --hatalar hatalar.xlsx --birim inc --top-esik 28 --parti-esik 20
```

**Metrik birim (varsayılan):** Top uzunluğu ve hata konumu metre, en ve hata uzunluğu cm cinsindendir.
`--birim inc` ile yard ve inç kullanılır. Kontrolcü puanı kendisi verdiyse `Puan` sütunu doğrudan kullanılır.

## Çıktı

`Partiler` · `Toplar` (grafikli) · `Hata Analizi` (Pareto) · `Hata Kayıtları` · `Bilgi`

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
