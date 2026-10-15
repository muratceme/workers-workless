# Rezervasyon Pick-up Raporu · Kod Bloğu

> Rezervasyon ve Gelir Yönetimi › Gelir Müdürü (Revenue Manager) · Workers / Workless

Gelecek konaklama tarihleri için **iki rezervasyon dökümü arasındaki artışı (pick-up)** hesaplar; geçen yıl ve
bütçeyle karşılaştırır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Pick-up (oda) | Bugünkü satılan oda − önceki dökümdeki satılan oda |
| Pick-up (gelir) | Bugünkü gelir − önceki gelir |
| ADR | Gelir ÷ satılan oda |
| Pick-up ADR | Pick-up gelir ÷ pick-up oda (yalnız pick-up oda pozitifse) |
| Doluluk % | Satılan oda ÷ kapasite × 100 |
| Geçen yıla fark | Satılan oda − geçen yıl satılan oda; yüzde olarak da verilir |
| Bütçe gerçekleşme % | Satılan oda ÷ bütçe oda × 100 |

- Aynı tarih dosyada birden çok satırdaysa (oda tipi, segment) toplanır.
- Aylık özet ve toplamda ADR ve oranlar toplamlardan hesaplanır; geçen yıl ve bütçe karşılaştırması yalnız o verisi
  bulunan günler üzerinden yapılır.
- Negatif pick-up iptal, grup düşüşü veya tarih değişikliği demektir; ayrıca işaretlenir.

**Örnek:** 23.10 için önceki döküm 96 oda / 300.000, bugünkü döküm 110 oda / 352.000. Pick-up 14 oda, 52.000 gelir;
pick-up ADR 3.714,29; kapasite 120 ise doluluk %91,67.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, kapasite 120
python main.py --bugun bugun.xlsx --onceki onceki.xlsx --kapasite 120
python main.py --bugun bugun.xlsx --onceki onceki.xlsx --gecen-yil gecen_yil.xlsx --butce butce.xlsx --kapasite 120
```

Bütün dosyalar aynı biçimdedir: Konaklama Tarihi, Satılan Oda, Gelir. "Önceki" dosya, PMS'ten daha önce (örneğin
7 gün önce) aldığınız aynı rapordur.

## Çıktı

`cikti/pick_up_raporu.xlsx`:

- `Günlük Pick-up`: tarih bazında satılan oda, pick-up, ADR, doluluk, geçen yıl ve bütçe karşılaştırması.
- `Aylık Özet`: ay ve toplam bazında aynı göstergeler.
- `Uyarılar`: negatif pick-up, kapasite üstü satış, bir dökümde eksik tarih, okunamayan satırlar.

## Dikkat

- **Dökümler aynı esasta olmalı.** Aynı oda tipleri, aynı rezervasyon durumları (kesin, opsiyonlu) ve aynı gelir
  tanımı (KDV, kahvaltı, konaklama vergisi dahil veya hariç). Tanım değişirse pick-up anlamını yitirir.
- **Geçen yılı siz hizalarsınız.** Dosyadaki tarihler bu yılın tarihleri olmalıdır. Aynı tarihi mi, aynı hafta gününü
  mü, bayram kaymasını mı esas alacağınıza siz karar verirsiniz; kod yalnız tarih eşleştirir.
- **Geçen yıl verisi "aynı gün kala" olmalı.** Bugünkü dökümü geçen yılın gerçekleşen sonucuyla karşılaştırmak hız
  hakkında yanıltır; mümkünse geçen yılın aynı kalan gündeki dökümünü kullanın.
- **Pick-up tahmin değildir.** Geçmiş hızın süreceği varsayılmaz; rapor ne olduğunu gösterir.
- **Kapasite sabit alınır.** Arızalı veya satışa kapalı odalar için kapasiteyi siz düzeltirsiniz.
- **Örnek veri:** Tarihler, oda sayıları ve gelirler kurgusaldır.

## Testler

- Günlük pick-up, ADR, pick-up ADR, doluluk; negatif pick-up, kapasite üstü, bir dökümde olmayan tarih.
- Geçen yıl ve bütçe karşılaştırması; aylık özet ve toplamlar; isteğe bağlı veriler olmadan çalışma.
- Aynı tarihli satırların toplanması, okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

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
