# Navlun Teklif Hesaplama · Kod Bloğu

> Deniz ve Hava Yolu Operasyon › Denizyolu Operasyon Uzmanı · Workers / Workless

Ölçü ve ağırlıktan **hava kargoda ücrete esas ağırlığı (1:6000), denizyolunda W/M'yi** hesaplar; tarife dosyanızdaki
navlun ve ek masraflarla teklif oluşturur. İnternete bağlanmaz.

## Nasıl hesaplar?

| Taşıma | Formül |
|---|---|
| Hava | Hacimsel ağırlık (kg) = Σ adet × boy × en × yükseklik (cm) ÷ 6000. Ücrete esas ağırlık = en büyük(brüt, hacimsel), 0,5 kg'a yukarı yuvarlanır |
| Deniz (LCL) | Hacim (m³) = Σ adet × boy × en × yükseklik ÷ 1.000.000; ton = brüt kg ÷ 1000. W/M = en büyük(ton, m³) |

| Tarife birimi | Tutar |
|---|---|
| Ücrete esas (kg veya W/M) | Birim fiyat × ücrete esas ağırlık veya W/M |
| Kap | Birim fiyat × kap adedi |
| Sabit | Birim fiyat |

"Asgari" yazılmışsa kalemin tutarı onun altına inmez. Toplamlar para birimi bazında ayrı verilir.

**Örnek (hava):** 2 koli 60 × 50 × 40 cm, 25 kg ve 1 sandık 120 × 80 × 110 cm, 150 kg → hacim 1.296.000 cm³ ÷ 6000 = 216 kg
hacimsel; brüt 200 kg → ücrete esas 216,0 kg. Navlun 2,50 × 216 = 540,00.
**Aynı yük denizde:** 1,296 m³, 0,200 ton → W/M 1,296.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri, hava
python main.py --mod deniz                             # örnek veri, deniz
python main.py --kaplar kaplar.xlsx --tarife tarife.xlsx --mod hava --bolen 6000 --yuvarlama 0,5
```

| Dosya | Sütunlar |
|---|---|
| Kaplar | Adet, Boy (cm), En (cm), Yükseklik (cm), Ağırlık (kg, kap başına) (zorunlu); Kap |
| Tarife | Kalem, Birim Fiyat (zorunlu); Birim (Ücrete esas / W/M / Kap / Sabit), Asgari, Para Birimi |

## Çıktı

`cikti/navlun_teklifi.xlsx`:

- `Teklif`: masraf kalemleri, çarpan, asgari uygulaması, para birimi bazında toplam; hesap özeti.
- `Kap Listesi`: kap bazında brüt ağırlık, hacim ve (havada) hacimsel ağırlık.
- `Uyarılar`: okunamayan satırlar, birden çok para birimi.

## Dikkat

- **Fiyatlar sizin tarifenizdir.** Kod hiçbir havayolu, hat veya acentenin fiyatını bilmez; örnek dosyalardaki
  rakamlar "örnek fiyat" olarak işaretlidir.
- **Bölen ve yuvarlama taşıyıcıya göre değişir.** 1:6000 hava kargoda yaygın orandır; ekspres taşıyıcılar ve bazı
  hatlar farklı bölen ve yuvarlama kullanır. `--bolen` ve `--yuvarlama` ile ayarlayın.
- **W/M parsiyel (LCL) içindir.** Tam konteyner (FCL) navlunu konteyner başınadır; tarifede "Sabit" kalem olarak
  girin. 1 ton = 1 m³ kabulü dışında oran kullanan tarifeler desteklenmez.
- **Ağırlık kademeleri yoktur.** Havada ağırlık dilimine göre değişen birim fiyat (45 kg, 100 kg üstü gibi) ve bir üst
  dilimden ucuza gelme kontrolü yapılmaz; doğru dilimin fiyatını tarifeye siz yazın.
- **Ölçüm taşıyıcıya aittir.** Kap ölçüleri en dış noktadan alınır; taşıyıcının tartı ve ölçümü esas olur.
- **Örnek veri:** Kaplar ve fiyatlar kurgusaldır.

## Testler

- Hava: hacimsel ağırlık, 0,5 kg yuvarlama sınırları, ağırlığın belirlediği durum, farklı bölen.
- Deniz: W/M, asgari tutar, birden çok para birimi; tarife birimi tanıma.
- Okuma hataları ve iki örnek tarife uçtan uca.

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
