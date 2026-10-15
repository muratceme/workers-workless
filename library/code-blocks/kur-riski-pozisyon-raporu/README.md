# Kur Riski Pozisyon Raporu · Kod Bloğu

> Finans › Finans Müdürü · Workers / Workless

Döviz cinsinden varlık ve yükümlülüklerinizden **net döviz pozisyonunu** çıkarır ve "kur %10 artarsa ne olur?"
sorusunu TL olarak yanıtlar. Para birimi ve vade dilimi bazında açığın nerede olduğunu gösterir. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Bilanço net pozisyonu | Döviz varlıklar − döviz yükümlülükler (her para birimi ayrı) |
| Türev net | Türev alım − türev satım (vadeli döviz alım/satım gibi) |
| Net pozisyon | Bilanço net + türev net. Artı = uzun (döviz fazlası), eksi = kısa (döviz açığı) |
| TL karşılığı | Net pozisyon × kur |
| Şok etkisi | Net pozisyon (TL) × şok %. Kur artınca açık pozisyon zarar, fazla pozisyon kâr yazar |
| Pozisyon / özkaynak | Toplam net pozisyon (TL) ÷ `--ozkaynak` |
| Vade kırılımı | Vadesi yazılan kalemler 0–30, 31–90, 91–365 gün ve 1 yıldan uzun dilimlerine ayrılır |

**Örnek:** EUR varlık 400.000, yükümlülük 1.650.000, vadeli alım 500.000 → net −750.000 EUR. Kur 49 ise
−36.750.000 TL; kur %10 artarsa kur farkı zararı **3.675.000 TL**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                        # örnek veri, durum tarihi 12.10.2026
python main.py --kalemler kalemler.xlsx --kurlar kurlar.xlsx
python main.py --kalemler kalemler.xlsx --kurlar kurlar.xlsx --ozkaynak 60.000.000 --limit 30 "--sok=5;15;-5"
```

Zorunlu sütunlar: Tür, Para Birimi, Tutar. Diğerleri: Kalem, Vade. Tutarları **artı işaretli** yazın; yönü Tür
sütunu belirler. Kendi dosyanızla çalışırken `--kurlar` zorunludur.

| Seçenek | Anlamı | Varsayılan |
|---|---|---|
| `--sok` | Kur şoku senaryoları, yüzde; noktalı virgülle ayrılır | 10;20;-10 |
| `--ozkaynak` | Özkaynak tutarı; verilirse pozisyon / özkaynak oranı hesaplanır | yok |
| `--limit` | Bu oran için kendi belirlediğiniz sınır; aşılırsa uyarılır | yok |

## Çıktı

`cikti/kur_riski_pozisyon.xlsx`:

- `Pozisyon`: para birimi bazında varlık, yükümlülük, türev, net pozisyon, yön ve TL karşılığı.
- `Senaryolar`: her kur şokunda para birimi ve toplam bazında kur farkı etkisi.
- `Vade Kırılımı`: para birimi ve vade dilimi bazında net tutar; açığın hangi vadede olduğunu gösterir.
- `Kalemler`, `Uyarılar`: girilen kalemler; eksik kur, riski büyüten türev, sınır aşımı, vadesi geçmiş kalemler.

## Dikkat

- **Bu bir bakiye fotoğrafıdır.** Gelecekteki dövizli satışlar, alımlar ve faiz ödemeleri (nakit akış riski)
  hesaba girmez; ihracat geliri olan bir şirketin gerçek riski tablodakinden düşük olabilir.
- **Senaryolar basittir.** Bütün kurların TL karşısında aynı oranda değiştiği varsayılır; para birimleri arası
  çapraz kur hareketi, opsiyonların doğrusal olmayan etkisi ve vergi etkisi yoktur.
- **Kurları ve kapsamı siz belirlersiniz.** Hangi hesapların parasal kalem sayılacağı, dövize endeksli TL
  kalemlerin ve stokların dahil edilip edilmeyeceği sizin kararınızdır. Örnek kurlar kurgusaldır.
- **Sınır sizin politikanızdır.** `--limit` yasal bir oran değildir; şirketinizin risk politikasındaki değeri
  girin. Kredi sözleşmelerinizdeki şartları ve sektörünüze özgü düzenlemeleri ayrıca izleyin.
- **Örnek veri:** Kalemler, tutarlar ve kurlar kurgusaldır.

## Testler

- Para birimi bazında net pozisyon, türev etkisi, TL karşılığı ve şok senaryoları (elle hesaplanmış).
- Vade dilimi sınırları (30, 31, 90, 91, 365, 366 gün), vadesi geçmiş kalem, özkaynak oranı ve sınır.
- Eksik kur, TL kalem, okuma hataları ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
