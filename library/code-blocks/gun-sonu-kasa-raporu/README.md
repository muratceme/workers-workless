# Gün Sonu Kasa Raporu · Kod Bloğu

> Mağaza Operasyonları › Kasa Sorumlusu · Workers / Workless

Kasa bazında Z raporu ve POS gün sonu toplamlarından **nakit, kredi kartı, iade ve iptal toplamlarını** çıkarır;
kasada olması gereken nakdi hesaplayıp **sayımla farkını** ve kart tutarının POS gün sonuyla uyumunu raporlar.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Formül |
|---|---|
| Beklenen nakit | Açılış nakdi + nakit satış − nakit iade − kasadan çıkış |
| Nakit farkı | Sayılan nakit − beklenen nakit (eksi: kasa eksiği, artı: kasa fazlası) |
| Net kart satışı | Kredi kartı satış − kart iade |
| Kart farkı | POS gün sonu toplamı − net kart satışı |
| Brüt satış | Nakit + kredi kartı + diğer ödemeler |
| Net satış | Brüt satış − iadeler |
| İade %, iptal % | İade ya da iptal tutarı ÷ brüt satış × 100 |
| Teslim edilecek nakit | Sayılan nakit − açılış nakdi |

Fark `--tolerans` (5) içindeyse kasa "Tamam" sayılır; değilse "Nakit eksiği", "Nakit fazlası" ya da "Kart farkı"
yazılır. Farkın mutlak değeri `--buyuk` (500) ve üstündeyse uyarı yüksek önemlidir.

**Örnek:** Açılış 1.000, nakit satış 12.980, nakit iade yok, kasadan çıkış 5.000 → beklenen nakit **8.980**.
Sayılan 8.830 → nakit farkı **−150** (kasa eksiği).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek veri
python main.py --kasalar kasa_gun_sonu.xlsx
python main.py --kasalar kasa_gun_sonu.xlsx --tolerans 10 --buyuk 250
```

Zorunlu sütunlar: Kasa, Nakit Satış. Her satır bir kasanın (ya da kasiyer vardiyasının) gün sonudur. Boş bırakılan
tutarlar sıfır sayılır; `Sayılan Nakit` ya da `POS Gün Sonu` boşsa ilgili fark hesaplanmaz.

## Çıktı

`cikti/gun_sonu_kasa_raporu.xlsx`:

- `Kasa Mutabakatı`: kasa başına beklenen ve sayılan nakit, nakit farkı, net kart satışı, POS gün sonu, kart farkı,
  teslim edilecek nakit; son satır mağaza toplamı.
- `Satış Özeti`: ödeme türlerine göre satış, iadeler, net satış, iade ve iptal oranları, fiş adedi.
- `Uyarılar`.

## Dikkat

- **Tutarları siz girersiniz.** Kod yazar kasaya, POS cihazına ya da bankaya bağlanmaz; Z raporundaki ve POS gün sonu
  fişindeki rakamların doğru aktarılması sizin sorumluluğunuzdadır.
- **İptaller beklenen nakdi değiştirmez.** İptal edilen fiş satış toplamına girmediği için nakit hesabında yer
  almaz; yalnız izleme amacıyla raporlanır. Z raporunuz iptalleri satışın içinde gösteriyorsa satış tutarını
  iptalden arındırarak yazın.
- **"Kasadan çıkış" bütün nakit çıkışlarıdır.** Gün içinde bankaya ya da merkez kasaya teslim edilen para, kasadan
  yapılan ödemeler ve bozukluk alışverişi buraya yazılır; yazılmazsa kasa eksiği görünür.
- **"Diğer ödeme" mutabakatı yapılmaz.** Yemek kartı, hediye çeki, puan ve havale ile ödemeler yalnız satış
  toplamına girer; sağlayıcı raporlarıyla karşılaştırılmaz.
- **Teslim edilecek nakit bir varsayıma dayanır.** Ertesi gün için kasada açılış nakdi kadar tutar bırakıldığı
  kabul edilir; farklı bir tutar bırakıyorsanız rakamı buna göre düzeltin.
- **Kapsam dışı:** muhasebe kaydı, mali rapor (Z) numara takibi, taksitli satış ve komisyon, döviz ile tahsilat.
- **Örnek veri:** Kasalar, kasiyerler ve tutarlar kurgusaldır.

## Testler

- Beklenen nakit, nakit ve kart farkı, tolerans sınırı, büyük fark eşiği.
- Sayımı ya da POS gün sonu girilmemiş kasa, iade ve iptal oranları, mağaza toplamı.
- Mükerrer kasa, eksi tutar, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
