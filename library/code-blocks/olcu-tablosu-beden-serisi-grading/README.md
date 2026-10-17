# Ölçü Tablosu Beden Serisi (Grading) · Kod Bloğu

> Tekstil › Ürün Geliştirme / Modelhane · Gerisi Bende

Ana bedenin (numune bedeni) ölçülerinden ve ölçü noktası bazındaki **beden artış kurallarından** tüm beden
serisinin ölçü tablosunu üretir. İnternete bağlanmaz.

- **Artışlar:** Her beden geçişi için ayrı verilebilir (`XS-S`, `S-M`, `M-L` ...). Örneğin büyük bedenlerde
  artış genişleyebilir. Ya da tek bir `Artış` sütunuyla tüm geçişlere aynı değer uygulanır.
- **Hesap:** Ana bedenin üstündeki bedenlere artış eklenir, altındakilerden çıkarılır.
- **Beden serisi:** Harf (`XS S M L XL XXL`) veya numara (`36 38 40 42 44 46`) olabilir.
- **Kontroller:**
  - Eksik geçiş artışı (0 kabul edilir ve uyarılır)
  - Sıfır veya negatif ölçü
  - Büyük bedenin küçük bedenden küçük çıkması
  - Artışın toleranstan **küçük** olması (komşu bedenler ölçüyle ayırt edilemez)
- **Yuvarlama:** Varsayılan 0,1 cm; `--adim 0.5` ile değiştirilir.
- **İnç:** `--inc` ile inç karşılıkları ayrı sayfada verilir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                     # örnek erkek tişört (XS–XXL, ana M)
python main.py --girdi kurallar.xlsx --bedenler XS S M L XL XXL --ana M
python main.py --girdi kurallar.xlsx --bedenler 36 38 40 42 44 46 --ana 40 --inc --adim 0.5
```

> Örnek dosyadaki ölçü ve artışlar **örnektir**. Marka/alıcı ölçü tablonuzu ve grading kurallarınızı kullanın.

## Çıktı

`Ölçü Tablosu (cm)` (ana beden vurgulu, kontrol sütunu) · `Artış Kuralları` · `Ölçü Tablosu (inç)` · `Bilgi`

## Testler

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
