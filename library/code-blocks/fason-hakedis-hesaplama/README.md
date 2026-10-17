# Fason Hakediş Hesaplama · Kod Bloğu

> Üretim ve Fason Takip › Fason Takip Sorumlusu · Gerisi Bende

Atölyeden dönen **sağlam adetler, birim fiyat, eksik ve defolu ürün kesintilerinden** atölye bazında fason hakedişini
ve net ödenecek tutarı hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Eksik adet        = verilen − sağlam − defolu
Brüt hakediş      = sağlam adet × fason birim fiyatı
Tolerans adedi    = verilen × fire toleransı %
Kesintiye tabi    = eksik + defolu − tolerans adedi          (eksiyse 0)
Malzeme kesintisi = kesintiye tabi adet × malzeme bedeli
KDV               = brüt hakediş × KDV %
Tevkif edilen KDV = KDV × tevkifat oranı
Net ödenecek      = brüt hakediş + KDV − tevkif edilen KDV − malzeme kesintisi − diğer kesinti − avans
```

- Defolu ve eksik ürüne işçilik ödenmez; yalnız sağlam adet hakedişe girer.
- **Malzeme bedeli**, kaybolan veya kullanılamaz hale gelen bir ürünün kumaş ve aksesuar maliyetidir (satırda girilir).
- **Örnek:** 1.025 verildi, 1.000 sağlam, 15 defolu → eksik 10. Tolerans %1 = 10,25 adet → kesintiye tabi 14,75 adet
  × 180 TL = 2.655 TL. Brüt hakediş 1.000 × 40 = 40.000 TL.
- Defolu oranı `--defolu-esik` (varsayılan %3) üstündeki satırlar işaretlenir.
- Sağlam + defolu verilenden büyükse eksik 0 sayılır ve uyarı verilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (tolerans %1, KDV %10, tevkifat 7/10 ile)
python main.py --teslimler fason_teslimler.xlsx --tolerans 2 --kdv 10 --tevkifat 7/10
python main.py --teslimler fason_teslimler.xlsx        # tolerans, KDV ve tevkifat 0
```

Sütunlar: Atölye, Verilen Adet, Sağlam Adet, Birim Fiyat (zorunlu); Sipariş No, Model, Defolu Adet, Malzeme Bedeli,
Avans, Diğer Kesinti. Kendi dosyanızda `--tolerans`, `--kdv` ve `--tevkifat` verilmezse 0 kabul edilir.

## Çıktı

`cikti/fason_hakedis.xlsx`:

- `Hakediş Özeti`: atölye bazında brüt hakediş, KDV, tevkifat, kesintiler, avans ve net ödenecek; boş "Mutabakat" sütunu.
- `Satırlar`: sipariş / model bazında eksik, defolu, tolerans adedi, kesintiye tabi adet ve kesinti.
- `Uyarılar`: verilenden fazla teslim, malzeme bedeli olmayan kesinti, eksiye düşen net ödeme, okunamayan satırlar.

## Dikkat

- **KDV ve tevkifat oranlarını siz verirsiniz.** Fason işlerde KDV oranı ve tevkifat uygulaması mevzuata ve işin
  türüne göre değişir; güncel oranları mali müşavirinizle doğrulayın. Örnek verideki oranlar yalnız hesabı göstermek içindir.
- **Ön mutabakattır, fatura değildir.** Malzeme kesintisinin ayrı fatura (yansıtma) ile mi yoksa mahsupla mı
  yapılacağı ve bunun KDV'si muhasebe uygulamanıza bağlıdır. Bu hesapta KDV matrahı brüt hakediştir.
- **Tolerans ve kesinti koşulları sözleşmeden gelir.** Fire toleransı, defolu ürünün kimde kalacağı ve malzeme bedeli
  fason sözleşmenizde yazılı olmalıdır; kod bunları yorumlamaz.
- **Tamirden dönen ürünler:** Tamire verilen ürünleri sağlam olarak teslim alındıklarında sağlam adede ekleyin.
- **Verilen adet** atölyeye fiilen teslim edilen kesim adedidir; sipariş adedi değildir.
- **Örnek veri:** Atölyeler, adetler ve fiyatlar kurgusaldır.

## Testler

- Eksik, tolerans, kesintiye tabi adet ve malzeme kesintisi (toleranslı / toleranssız / tolerans altında).
- Atölye bazında KDV, tevkifat ve net ödeme.
- Fazla teslim, malzeme bedeli olmayan kesinti, eksi net ödeme; okuma hataları; örnek veri uçtan uca.

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
