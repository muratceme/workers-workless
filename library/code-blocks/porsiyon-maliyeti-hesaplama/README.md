# Porsiyon Maliyeti Hesaplama · Kod Bloğu

> Yiyecek ve İçecek › Yiyecek İçecek Müdürü · Workers / Workless

Standart reçete, gramaj ve güncel alış fiyatlarından menü kalemi başına **porsiyon maliyetini, maliyet oranını ve
hedef satış fiyatını** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Brüt miktar | Reçetedeki net miktar ÷ (1 − fire % ÷ 100) |
| Malzeme maliyeti | Brüt miktar (alış birimine çevrilmiş) × alış fiyatı |
| Porsiyon maliyeti | Σ malzeme maliyeti ÷ reçetenin porsiyon sayısı |
| Maliyet oranı % | Porsiyon maliyeti ÷ satış fiyatı (KDV hariç) × 100 |
| Brüt kâr | Satış fiyatı (KDV hariç) − porsiyon maliyeti |
| Hedef fiyat | Porsiyon maliyeti ÷ (`--hedef` % ÷ 100), KDV dahil |

Birimler: `g`, `kg` · `ml`, `cl`, `l` · `adet`. Gram ile kilogram, mililitre ile litre kendi aralarında çevrilir.
Ağırlık ile hacim veya adet arasında çeviri yapılmaz; böyle satırlar "birim uyuşmuyor" olarak işaretlenir.

**Örnek:** 150 g kıyma (600 TL/kg) = 90,00; 30 g soğan, fire %10 → 33,33 g (40 TL/kg) = 1,33; 1 ekmek = 5,00; 10 ml
yağ (200 TL/l) = 2,00. Porsiyon maliyeti 98,33. Satış fiyatı 320 → maliyet oranı %30,73; hedef %30 için fiyat 327,78.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --receteler receteler.xlsx --fiyatlar fiyatlar.xlsx
python main.py --receteler receteler.xlsx --fiyatlar fiyatlar.xlsx --menu menu.xlsx --hedef 28 --kdv 10
```

- Reçeteler: Menü Kalemi, Malzeme, Miktar, Birim, Fire % (her malzeme bir satır).
- Fiyatlar: Malzeme, Alış Fiyatı, Birim (fiyatın ait olduğu birim: kg, l, adet).
- Menü (isteğe bağlı): Menü Kalemi, Satış Fiyatı, Porsiyon Sayısı (reçete kaç porsiyon çıkarıyor; boşsa 1).

Satış fiyatları KDV dahilse `--kdv` ile oranı verin. Miktarlarda binlik ayırıcı kullanmayın (`1250,5`).

## Çıktı

`cikti/porsiyon_maliyeti.xlsx`:

- `Porsiyon Maliyeti`: reçete ve porsiyon maliyeti, maliyet oranı, brüt kâr, hedef fiyat, durum.
- `Malzeme Detayı`: brüt miktar, alış fiyatı, maliyet ve maliyet payı.
- `Uyarılar`: fiyatı olmayan veya birimi uyuşmayan malzemeler, hedefi aşan kalemler, okunamayan satırlar.

## Dikkat

- **Yalnız malzeme maliyetidir.** İşçilik, enerji, kira, servis ve zayi dahil değildir. Maliyet oranı tek başına
  kârlılığı göstermez; brüt kâr tutarına da bakın.
- **Hedef oran bir varsayımdır.** %30 yaygın bir başlangıç değeridir ama konsepte, ürüne ve fiyat düzeyine göre
  değişir; kanıtlanmış "doğru" oran değildir.
- **Fire sizin ölçümünüzdür.** Temizleme ve pişirme kaybı ürüne, tedarikçiye ve mevsime göre değişir; verim testi
  yaparak güncelleyin.
- **Adetle alınan ürünleri adetle yazın.** Reçetede "2 adet limon" varken fiyat kilogramla girilmişse kod çeviremez;
  ortalama ağırlıkla grama çevirin veya adet fiyatı girin.
- **Garnitür, sos, ekmek ve ikramları reçeteye ekleyin.** Tabağa giren ama reçetede olmayan her şey maliyeti düşük gösterir.
- **Alış fiyatının KDV'li mi KDV'siz mi olduğunu siz belirlersiniz;** tüm fiyatları aynı esasta girin.
- **Örnek veri:** Reçeteler, fiyatlar ve menü kurgusaldır.

## Testler

- Brüt miktar, birim çevirme (g, kg, ml, cl, l, adet), porsiyon sayısı, maliyet oranı, brüt kâr, hedef fiyat, KDV.
- Fiyatı olmayan ve birimi uyuşmayan malzeme, satış fiyatı olmayan kalem.
- Okunamayan reçete, fiyat ve menü satırları, komut satırı hataları; örnek veri uçtan uca.

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
