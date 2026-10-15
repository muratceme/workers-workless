# Ölçüm Verisi Uygunluk Kontrolü · Kod Bloğu

> Kalite Güvence ve Kalite Kontrol › Kalite Kontrol Elemanı · Workers / Workless

Ölçüm cihazı çıktılarını **teknik resim toleranslarıyla** karşılaştırır; uygunsuz ve sınıra yakın ölçüleri işaretler,
parça bazında KABUL / RET sonucu verir. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Alt limit = nominal + alt tolerans      (işaretiyle: -0,021)
Üst limit = nominal + üst tolerans
```

`Alt Limit` / `Üst Limit` sütunları doluysa doğrudan onlar kullanılır. Tek taraflı toleransta yalnız biri girilir
(ör. Ra en çok 1,6 → yalnız Üst Limit; sertlik en az 58 HRC → yalnız Alt Limit).

| Sonuç | Koşul |
|---|---|
| Uygunsuz (üst) / (alt) | Ölçülen > üst limit / ölçülen < alt limit |
| Sınıra yakın | Uygun, ama en yakın limite uzaklığı tolerans aralığının `--sinir` (varsayılan %10) kadarı veya daha azı |
| Uygun | Diğerleri |

- **Limit değerinin kendisi uygundur** (25,000 üst limitli ölçüde 25,000 uygun, 25,001 uygunsuz).
- "Sınıra yakın" yalnız çift taraflı toleransta hesaplanır. Örnek: 24,979 – 25,000 aralığı 0,021; %10'u 0,0021 →
  24,980 sınıra yakındır.
- **Parça sonucu:** ölçülerinden biri uygunsuzsa RET.
- **Sayı yazımı:** virgül de nokta da ondalık ayırıcıdır; `1.250` bin iki yüz elli değil, 1,25'tir. Binlik ayırıcı kullanmayın.
- Parça kodu boş bırakılan tolerans satırı, o karakteristiğin geçtiği tüm parçalara uygulanır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --olcumler olcumler.xlsx --toleranslar toleranslar.xlsx --sinir 20
```

| Dosya | Sütunlar |
|---|---|
| Ölçümler | Karakteristik, Ölçülen (zorunlu); Parça Kodu, Numune No, Tarih, Operatör, Cihaz |
| Toleranslar | Karakteristik (zorunlu); Parça Kodu, Nominal, Alt Tolerans, Üst Tolerans, Alt Limit, Üst Limit, Birim |

Her ölçüm satırı tek bir ölçüdür (numune × karakteristik).

## Çıktı

`cikti/olcum_uygunluk.xlsx`:

- `Parça Özeti`: numune bazında KABUL / RET, uygunsuz ve sınıra yakın ölçüler; boş "Karar / Not" sütunu.
- `Karakteristik Özeti`: ölçü bazında en küçük, en büyük, ortalama, uygunsuz sayısı ve oranı.
- `Ölçüm Sonuçları`: her ölçüm için limitler, nominalden sapma, sonuç ve limit dışı miktar.
- `Uyarılar`: toleransı tanımsız ölçümler, hiç ölçülmemiş karakteristikler, okunamayan değerler.

## Dikkat

- **Ölçüm belirsizliği hesaba katılmaz.** Limite çok yakın sonuçlarda cihazın belirsizliği kararı değiştirebilir;
  "Sınıra yakın" ölçüleri bu gözle değerlendirin. Karar kuralınız (ör. belirsizlik payı düşülerek kabul) farklıysa
  limitleri ona göre girin.
- **Alt toleransı işaretiyle yazın.** ±0,05 için alt tolerans `-0,05`, üst tolerans `0,05` olmalıdır. İkisi de aynı
  artı değerse aralık sıfır çıkar ve uyarı verilir.
- **Geometrik toleranslar** (konum, salgı, düzlemsellik vb.) tek taraflı üst limit olarak girilebilir; kod bunların
  ölçüm yöntemini veya referans (datum) düzenini bilmez.
- **İstatistiksel yeterlilik hesaplamaz.** Cp / Cpk için ayrı bir paket (SPC kontrol grafiği) kullanın.
- **Nihai karar kalite sorumlusundadır.** Sapma onayı, yeniden işleme veya hurda kararı bu çıktıyla verilmez.
- **Örnek veri:** Parça, toleranslar ve ölçümler kurgusaldır.

## Testler

- Sayı okuma (virgül, nokta, eksi işareti, birimli değer).
- Çift taraflı tolerans: limit değerleri, taşma miktarı, sınıra yakın eşiği; tek taraflı tolerans.
- Parça ve karakteristik özetleri; toleransı tanımsız ölçüm, ölçülmemiş karakteristik.
- Okuma hataları ve örnek veri uçtan uca.

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
