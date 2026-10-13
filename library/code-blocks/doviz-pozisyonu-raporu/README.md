# Döviz Pozisyonu Raporu · Kod Bloğu

> Hazine ve Aktif-Pasif Yönetimi › Hazine Uzmanı · Workers / Workless

Bilanço içi döviz varlık ve yükümlülükleri ile bilanço dışı alım-satım taahhütlerinden **döviz cinsi bazında net
pozisyonu** hesaplar; TL karşılığını, uzun ve kısa toplamı ve bir kur şokunun etkisini gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Formül |
|---|---|
| Bilanço içi net | Döviz varlıkları − döviz yükümlülükleri |
| Bilanço dışı net | Döviz alım taahhütleri − döviz satım taahhütleri |
| Net genel pozisyon | Bilanço içi net + bilanço dışı net (artı: uzun, eksi: kısa) |
| TL karşılığı | Pozisyon × kur |
| Uzun / kısa toplam | TL karşılığı artı olan ve eksi olan pozisyonların ayrı ayrı toplamı |
| Pozisyon ÷ özkaynak % | Net genel pozisyon toplamı (TL) ÷ özkaynak × 100 |
| Kur şoku etkisi | Pozisyonun TL karşılığı × `--sok` (10) ÷ 100 |

**Örnek:** USD varlık 90 milyon, yükümlülük 100 milyon → bilanço içi **−10 milyon** (kısa). Swap ve vadeli işlemlerde
14 milyon alım, 2 milyon satım → bilanço dışı **+12 milyon**. Net genel pozisyon **+2 milyon USD** (uzun); kur 42,50
→ 85.000.000 TL. Kur %10 artarsa kazanç 8.500.000 TL.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri ve örnek kurlar
python main.py --kalemler kalemler.xlsx --kurlar kurlar.xlsx
python main.py --kalemler kalemler.xlsx --kurlar kurlar.xlsx --ozkaynak 2000000000 --limit 20 --sok 15
```

Zorunlu sütunlar: Tür, Para Birimi, Tutar. Tutarlar orijinal para biriminde ve artı işaretli yazılır; yönü tür sütunu
belirler. TL kalemler dikkate alınmaz.

## Çıktı

`cikti/doviz_pozisyonu.xlsx`:

- `Pozisyon`: para birimi bazında varlık, yükümlülük, bilanço içi net, bilanço dışı alım / satım / net, net genel
  pozisyon, yön, kur, TL karşılıkları ve kur şoku etkisi; altında uzun, kısa ve net toplam.
- `Kalemler`: hesaba giren kalemler (denetim izi).
- `Uyarılar`: kuru eksik para birimi, limit aşımı ya da limite yaklaşma.

## Dikkat

- **Bu yasal oran hesabı değildir.** Yabancı para net genel pozisyonunun özkaynağa oranına ilişkin düzenlemede
  hangi kalemlerin dahil edileceği, dövize endeksli varlık ve yükümlülükler, opsiyonların delta eşdeğeri, karşılıklar
  ve istisnalar ayrıca tanımlanır; kod yalnız yazdığınız dört türü toplar. Limit yüzdesini de siz verirsiniz; kod
  bir yasal sınır varsaymaz. Yasal raporlama için kullanmayın.
- **Bilanço dışı kalemleri nominal tutarla girin** ve bir işlemin iki bacağını (aldığınız ve sattığınız para
  birimi) ayrı satırlara yazın. USD alıp EUR satan bir swap, USD için "Bilanço Dışı Alım", EUR için "Bilanço Dışı
  Satım" satırıdır.
- **Kurlar sizin girdinizdir.** Örnek dosyadaki kurlar kurgusaldır. Değerleme kurunu bütün kalemlerde aynı tarihten
  alın.
- **Kur şoku doğrusaldır ve bütün para birimlerine aynı yüzdeyi uygular.** Çapraz kur hareketleri, opsiyonların
  doğrusal olmayan etkisi ve vade yapısı hesaba katılmaz.
- **Kapsam dışı:** vade dilimlerine göre döviz likiditesi, faiz riski, altın ve kıymetli maden hesaplarının özel
  durumu, riske maruz değer.
- **Örnek veri:** Kalemler, tutarlar ve kurlar kurgusaldır.

## Testler

- Bilanço içi, bilanço dışı ve net genel pozisyon; uzun, kısa ve kapalı pozisyon; TL karşılığı ve toplamlar.
- Özkaynağa oran, limit aşımı ve yaklaşma sınırları (kısa pozisyonda mutlak değer), kur şoku etkisi.
- Kuru olmayan para birimi, TL kalemler, tür yazımları, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
