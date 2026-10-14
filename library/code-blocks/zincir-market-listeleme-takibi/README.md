# Zincir Market Listeleme Takibi · Kod Bloğu

> Satış ve Kanal Yönetimi › Kanal Satış Yöneticisi · Workers / Workless

Saha ekibinin mağaza ziyaret kayıtlarından ürünlerin **zincir bazında listeleme oranını**, listeli olup rafta stoğu
olmayan (**stoksuz**) mağazaları ve raf fiyatlarının **önerilen raf fiyatından sapmasını** çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Güncel durum | Her zincir · mağaza · ürün için en son ziyaret kaydı |
| Listeleme oranı % | Ürünün listeli olduğu mağaza ÷ ziyaret edilen mağaza × 100 |
| Stoksuzluk % | Listeli olup rafta stok olmayan mağaza ÷ listeli mağaza × 100 |
| Fiyat sapması % | (Ortalama raf fiyatı − önerilen raf fiyatı) ÷ önerilen × 100 |

Listeli sütunu boşsa fiyat ya da stok yazılmış satır listeli sayılır. Stok sütunu "Var / Yok" ya da adet olabilir
(0 = yok).

**Uyarılar:** listeleme `--liste` (varsayılan %80) altında (listesiz mağazalar adıyla) · stoksuzluk `--oos`
(varsayılan %10) üstünde · ortalama raf fiyatı önerilenden ±`--fiyat` (varsayılan %5) fazla sapıyor · ürün listesinde
olup hiç taranmamış ürün.

**Örnek:** Zincir B'de yeşil zeytin 800 g ziyaret edilen 6 mağazanın hiçbirinde listeli değil → listeleme **%0**;
zincir genelinde listeli ürünlerin **%21,05**'i rafta stoksuz.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                  # örnek veri
python main.py --tarama magaza_ziyaretleri.xlsx --urunler urun_listesi.xlsx
python main.py --tarama magaza_ziyaretleri.xlsx --liste 90 --oos 5 --fiyat 3
```

- **Tarama:** Tarih, Zincir, Mağaza, Ürün Kodu, Listeli (Evet / Hayır), Raf Fiyatı, Stok Durumu.
- **Ürünler (isteğe bağlı):** Ürün Kodu, Ürün Adı, Önerilen Raf Fiyatı.

Bir mağazada bir ürün listeli değilse de satır yazılmalıdır (Listeli = Hayır); yazılmayan ürün o mağazada ziyaret
edilmemiş sayılır.

## Çıktı

`cikti/zincir_listeleme_takibi.xlsx`:

- `Zincir × Ürün`: ziyaret, listeli, listeleme %, stoksuz, stoksuzluk %, ortalama / en düşük / en yüksek raf fiyatı,
  önerilen fiyat, sapma %.
- `Zincirler`: zincir geneli listeleme ve stoksuzluk.
- `Mağaza Detayı`: mağaza × ürün güncel durumu.
- `Uyarılar`.

## Dikkat

- **Raf fiyatı zincirin kararıdır.** Önerilen raf fiyatı tavsiyedir; tedarikçinin perakendecinin satış fiyatını
  belirlemesi ya da zorlaması rekabet hukuku açısından yasaktır (yeniden satış fiyatının belirlenmesi). Fiyat sapması
  yalnız izleme amaçlıdır; zincirle fiyat görüşmesini hukuk biriminizin onayladığı çerçevede yapın.
- **Saha verisi anlıktır.** Stoksuzluk ziyaret anındaki raf durumudur; depoda stok olup rafa çıkmamış ürün de stoksuz
  görünür.
- **Eşikler politika varsayılanıdır;** listeleme anlaşmalarınıza göre değiştirin.
- **Örnek veri:** Zincirler, mağazalar, ürünler ve fiyatlar kurgusaldır.

## Testler

- En son ziyaretin seçimi, listeleme ve stoksuzluk oranı, ortalama fiyat ve sapma, zincir toplamları.
- Uyarı eşikleri, taranmamış ürün, listeli / stok yazımlarının tanınması, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
