# Ölçü Tablosu Beden Serisi (Grading) · Kod Bloğu

> Tekstil › Ürün Geliştirme / Modelhane · Workers / Workless

Ana bedenin (numune bedeni) ölçülerinden ve ölçü noktası bazındaki **beden artış kurallarından** tüm beden
serisinin ölçü tablosunu üretir. İnternete bağlanmaz.

- **Artışlar:** Her beden geçişi için ayrı verilebilir (`XS-S`, `S-M`, `M-L` ...). Örneğin büyük bedenlerde
  artış genişleyebilir. Ya da tek bir `Artış` sütunuyla tüm geçişlere aynı değer uygulanır.
- **Hesap:** Ana bedenin üstündeki bedenlere artış eklenir, altındakilerden çıkarılır.
- **Beden serisi:** Harf (`XS S M L XL XXL`) veya numara (`36 38 40 42 44 46`) olabilir.
- **Kontroller:**
  - Eksik geçiş artışı (0 kabul edilir ve uyarılır)
  - Sıfır veya negatif ölçü
  - Büyük bedenin küçük bedenden küçük çıkması
  - Artışın toleranstan **küçük** olması (komşu bedenler ölçüyle ayırt edilemez)
- **Yuvarlama:** Varsayılan 0,1 cm; `--adim 0.5` ile değiştirilir.
- **İnç:** `--inc` ile inç karşılıkları ayrı sayfada verilir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                     # örnek erkek tişört (XS–XXL, ana M)
python main.py --girdi kurallar.xlsx --bedenler XS S M L XL XXL --ana M
python main.py --girdi kurallar.xlsx --bedenler 36 38 40 42 44 46 --ana 40 --inc --adim 0.5
```

> Örnek dosyadaki ölçü ve artışlar **örnektir**. Marka/alıcı ölçü tablonuzu ve grading kurallarınızı kullanın.

## Çıktı

`Ölçü Tablosu (cm)` (ana beden vurgulu, kontrol sütunu) · `Artış Kuralları` · `Ölçü Tablosu (inç)` · `Bilgi`

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
