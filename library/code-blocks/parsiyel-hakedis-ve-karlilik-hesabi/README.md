# Parsiyel Hakediş ve Kârlılık Hesabı · Kod Bloğu

> Parsiyel Operasyon › Parsiyel Operasyon Uzmanı · Workers / Workless

Araçtaki yüklerin **navlun gelirlerini, araç maliyetini ve yük bazında kâr payını** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Yükün maliyet payı | Araç maliyeti × yükün ölçüsü ÷ araçtaki yüklerin toplam ölçüsü |
| Yük kârı, marj % | Navlun geliri − maliyet payı; kâr ÷ navlun geliri × 100 |
| Araç kârı | Araçtaki yüklerin navlun geliri toplamı − araç maliyeti |
| Doluluk % | Toplam LDM ÷ araç kapasitesi (dosyadaki "Kapasite LDM"; yoksa `--kapasite-ldm`, 13,6) |
| Dolu LDM başına maliyet | Araç maliyeti ÷ toplam LDM |

- **Dağıtım ölçüsü** `--anahtar` ile seçilir: `ldm` (varsayılan), `kg` veya `gelir`.
- Kuruş farkı aracın son yüküne eklenir; payların toplamı araç maliyetine tam eşittir.
- Araçtaki yüklerin seçilen ölçüsü toplamı 0 ise maliyet navlun gelirine göre dağıtılır ve uyarılır.

**Örnek:** araç maliyeti 3.000; yükler 4,0 / 3,2 / 4,8 LDM (toplam 12) ve 1.500 / 1.200 / 900 gelir → paylar 1.000 / 800
/ 1.200; kârlar +500 / +400 / −300. Araç kârı 3.600 − 3.000 = 600; doluluk 12 ÷ 13,6 = %88,24.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --yukler yukler.xlsx --seferler sefer_maliyetleri.xlsx --anahtar kg --kapasite-ldm 13,6
```

| Dosya | Sütunlar |
|---|---|
| Yükler | Sefer, Navlun Geliri (zorunlu); Müşteri, Yük Ref, LDM, Ağırlık (kg) |
| Sefer maliyetleri | Sefer, Araç Maliyeti (zorunlu); Kapasite LDM |

## Çıktı

`cikti/parsiyel_karlilik.xlsx`:

- `Araç Özeti`: araç başına gelir, maliyet (hakediş), kâr, doluluk, boş LDM, dolu LDM başına maliyet.
- `Yük Bazında`: pay oranı, maliyet payı, kâr ve marj; zarar eden yükler işaretli.
- `Müşteri Bazında`: müşteri toplamları.
- `Uyarılar`: maliyeti olmayan araçlar, ölçüsü olmayan yükler, %100'ü aşan doluluk.

## Dikkat

- **Yük bazındaki kâr anahtara bağlıdır.** LDM, ağırlık veya gelire göre dağıtım aynı yükü kârlı ya da zararlı
  gösterebilir; araç kârı ise değişmez. Fiyatlandırma kararını tek anahtara dayandırmayın.
- **Boş kapasitenin maliyeti yüklere biner.** Yarı dolu araçta her yük zararlı görünebilir; bu, yükün değil
  doluluğun sorunudur. `Araç Özeti`ndeki doluluğa birlikte bakın.
- **Araç maliyeti sizin verinizdir.** Alt nakliyeci hakedişi veya öz mal aracın sefer maliyeti olarak ne yazarsanız
  o kullanılır; toplama-dağıtım, depo aktarma, gümrük ve genel giderler dahil değilse kâr yüksek görünür.
- **Tek para birimi.** Kur çevirisi yapılmaz.
- **Örnek veri:** Seferler, müşteriler ve tutarlar kurgusaldır.

## Testler

- LDM, ağırlık ve gelire göre dağıtım; kuruş farkının son yüke eklenmesi.
- Ölçüsü olmayan yükler, maliyeti olmayan araç, araç ve müşteri özetleri, doluluk.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
