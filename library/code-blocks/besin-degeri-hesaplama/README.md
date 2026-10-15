# Besin Değeri Hesaplama · Kod Bloğu

> Kalite Güvence ve Gıda Güvenliği › Gıda Mühendisi · Workers / Workless

Reçetedeki bileşenlerin miktarlarını ve besin değeri tablosunu kullanarak son ürünün **100 g ve porsiyon başına**
enerji, yağ, doymuş yağ, karbonhidrat, şekerler, lif, protein ve tuz değerlerini, porsiyonda **referans alım
yüzdesini** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Besin öğesi / 100 g | Σ(bileşen miktarı × bileşenin 100 g'daki değeri ÷ 100) ÷ son ürün ağırlığı × 100 |
| Son ürün ağırlığı | `--son-agirlik`; verilmezse reçete toplamı × (1 − `--kayip` %) |
| Enerji | g başına: yağ 37 kJ / 9 kcal · karbonhidrat 17 / 4 · protein 17 / 4 · lif 8 / 2 · poliol 10 / 2,4 · alkol 29 / 7 · organik asit 13 / 3 |
| Tuz | Tuz sütunu; yoksa sodyum × 2,5 (Sodyum mg sütunu da okunur) |
| Referans alım % | Porsiyondaki miktar ÷ referans alım × 100 |

Enerji dönüşüm katsayıları, tuz = sodyum × 2,5 eşdeğerliği ve referans alım değerleri (enerji 8400 kJ / 2000 kcal, yağ
70 g, doymuş yağ 20 g, karbonhidrat 260 g, şekerler 90 g, protein 50 g, tuz 6 g) Türk Gıda Kodeksi Gıda Etiketleme ve
Tüketicileri Bilgilendirme Yönetmeliği'ndeki değerlerdir; güncel metinden doğrulayın.

Değerler 100 g için enerjide tam sayıya, tuzda iki, diğerlerinde bir ondalığa yuvarlanır. Etiket yuvarlama kuralları ayrıca
uygulanmalıdır.

**Örnek:** Kakaolu kurabiye reçetesi 1.084 g, pişmiş ürün 1.000 g → 100 g'da **446 kcal**, yağ 20,2 g, doymuş yağ 12,2 g,
şekerler 21,3 g, tuz 0,57 g. 40 g porsiyonda doymuş yağ referans alımın **%24,4**'ü.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                   # örnek veri (son ağırlık 1000 g, porsiyon 40 g)
python main.py --recete recete.xlsx --tablo besin_tablosu.xlsx --son-agirlik 1000 --porsiyon 40 --urun "Kakaolu kurabiye"
python main.py --recete recete.xlsx --tablo besin_tablosu.xlsx --kayip 8
```

- **Reçete:** Bileşen, Miktar (g).
- **Tablo:** Bileşen ve 100 g'daki Yağ, Doymuş Yağ, Karbonhidrat, Şeker, Lif, Protein, Tuz (ya da Sodyum); isteğe bağlı
  Poliol, Alkol, Organik Asit. Karbonhidrat değeri lif ve poliol **hariç** (kullanılabilir karbonhidrat) olmalıdır.

## Çıktı

`cikti/besin_degeri.xlsx`:

- `Besin Değeri Tablosu`: 100 g, porsiyon ve referans alım %.
- `Bileşen Katkısı`: her bileşenin reçetedeki payı ve besin öğelerine katkısı (g).
- `Uyarılar`: tabloda olmayan bileşen, tutarsız tablo değerleri, reçeteden ağır son ürün.

## Dikkat

- **Hesaplanmış değerdir.** Etikette beyan edilen besin değerleri laboratuvar analizi, hesaplama ya da genel kabul görmüş
  veriye dayanabilir; hesaplama kullanıyorsanız bileşen verisinin güvenilir kaynaktan (tedarikçi spesifikasyonu, ulusal
  gıda kompozisyon veri tabanı) gelmesi ve ürünün izin verilen toleranslar içinde kalması sizin sorumluluğunuzdadır.
- **İşleme kayıpları:** Kod yalnız su kaybını (ağırlık değişimini) hesaba katar; pişirmede vitamin ve yağ kayıpları ile
  kızartmada yağ emilimi hesaplanmaz.
- **Beslenme ve sağlık beyanları** (ör. "az yağlı", "lif kaynağı") için yönetmelikteki koşullar ayrıca kontrol edilmelidir.
- **Örnek veri:** Reçete ve tablo değerleri gösterim amaçlıdır; kendi bileşen verinizi kullanın.

## Testler

- 100 g ve porsiyon değerleri, enerji (kJ / kcal), tuz–sodyum çevrimi, referans alım %, bileşen katkısı.
- Ağırlık kaybı, eksik bileşen, tutarsız tablo değeri, okuma ve komut satırı.

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
