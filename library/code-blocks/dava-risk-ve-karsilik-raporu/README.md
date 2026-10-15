# Dava Risk ve Karşılık Raporu · Kod Bloğu

> Hukuk › Hukuk Müdürü · Workers / Workless

Devam eden dava listenizi, mali işlerin ve denetçinin dönem sonunda istediği tabloya çevirir: hangi dava için
**karşılık** ayrılacak, hangisi **koşullu yükümlülük** olarak dipnotta açıklanacak, önceki döneme göre karşılık ne
kadar artacak veya azalacak. İnternete bağlanmaz.

## Nasıl hesaplar?

Şirketin **davalı** olduğu davalar kaybetme olasılığına göre sınıflanır:

| Sınıf | Olasılık | Sonuç |
|---|---|---|
| Muhtemel | %50'den büyük | Karşılık ayrılır. Tutar: Tahmini Yükümlülük (boşsa Dava Değeri kullanılır ve uyarılır) |
| Mümkün | %10'dan büyük, %50 ve altı | Karşılık ayrılmaz; koşullu yükümlülük olarak dipnotta açıklanır |
| Uzak | %10 ve altı | Karşılık ayrılmaz, açıklama gerekmez |
| Değerlendirilmemiş | Olasılık boş | Karar verilemez; uyarılır |

Olasılığı yüzde yerine sözle de yazabilirsiniz: Muhtemel / Yüksek, Mümkün / Orta, Uzak / Düşük. Şirketin
**davacı** olduğu davalar "koşullu varlık" olarak ayrı listelenir; bunlar için karşılık ayrılmaz.

- **Dönem değişimi** = karşılık − Önceki Karşılık. Artı değer ek karşılık (dönem gideri), eksi değer konusu
  kalmayan karşılıktır.
- **Beklenen değer** = tutar × olasılık. Yalnız yönetim bilgisidir; karşılık tutarı değildir.

**Örnek:** Dava değeri 450.000, kaybetme olasılığı %70, avukat tahmini 380.000, önceki karşılık 300.000 →
Muhtemel; karşılık **380.000**, ek karşılık **80.000**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                  # örnek veri
python main.py --davalar davalar.xlsx
python main.py --davalar davalar.xlsx --muhtemel 50 --uzak 5
```

Zorunlu sütunlar: Dosya No, Dava Değeri. Diğerleri: Karşı Taraf, Konu, Taraf Sıfatı (Davalı / Davacı; boşsa
Davalı), Kaybetme Olasılığı, Tahmini Yükümlülük, Önceki Karşılık, Aşama, Avukat.

## Çıktı

`cikti/dava_risk_karsilik.xlsx`:

- `Dava Listesi`: sınıf, dava değeri, olasılık, karşılık, önceki karşılık, dönem değişimi, yapılacak işlem.
- `Özet`: sınıf bazında dosya sayısı ve tutar; önceki karşılık, ek karşılık, iptal ve dönem sonu karşılık.
- `Konu Kırılımı`, `Avukat Kırılımı`: dosya sayısı, dava değeri, karşılık ve koşullu yükümlülük.
- `Uyarılar`: değerlendirilmemiş davalar, tahmini olmayan muhtemel davalar, konusu kalmamış olabilecek karşılıklar.

## Dikkat

- **Bu bir çalışma tablosudur, muhasebe kararı değildir.** Sınıflama TMS 37 ve BOBİ FRS'deki genel yaklaşımı
  izler: çıkış olasılığı gerçekleşmeme olasılığından yüksekse karşılık, uzak değilse dipnot. Nihai karar mali
  işlerin ve bağımsız denetçinindir.
- **"Uzak" için standartta sayısal eşik yoktur.** %10 yaygın bir çalışma varsayımıdır; şirket politikanıza göre
  `--uzak` ile değiştirin ve dönemler arasında aynı tutun.
- **Olasılık ve tutar avukatın tahminidir.** Kod yalnız verilen değeri sınıflar. Tahmini yükümlülüğe faiz,
  yargılama gideri ve vekâlet ücretini dahil edin; kod bunları hesaplamaz.
- **Kapsam dışı:** paranın zaman değeri (iskonto), sigorta veya rücu yoluyla tazmin beklentisi, benzer çok
  sayıda davanın birlikte beklenen değerle ölçülmesi, vergi mevzuatında karşılığın gider yazılabilirliği.
- **Gizlilik:** Dava değerlendirmeleri avukat-müvekkil gizliliği kapsamındadır ve karşı tarafın eline geçmesi
  davaya zarar verebilir; raporu yalnız ilgili kişilerle paylaşın.
- **Örnek veri:** Dosyalar, taraflar ve tutarlar kurgusaldır.

## Testler

- Sınıf eşikleri (%50, %50,01, %10, %10,01), sözle verilen olasılık, davacı olunan dava.
- Karşılık tutarı, tahmin yokken dava değeri, dönem değişimi (ek ve iptal), beklenen değer, toplamlar.
- Konu ve avukat kırılımı, okuma hataları ve komut satırı.

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
