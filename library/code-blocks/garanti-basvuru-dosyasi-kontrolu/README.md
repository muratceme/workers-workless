# Garanti Başvuru Dosyası Kontrolü · Kod Bloğu

> Satış Sonrası Hizmetler (Servis) › Garanti Uzmanı · Workers / Workless

Garanti başvurularını üreticiye göndermeden önce **şasi, süre, km, kod ve parça bilgileri** yönünden denetler; ret
riski taşıyan ve düzeltilmesi gereken başvuruları işaretler. İnternete bağlanmaz.

## Nasıl kontrol eder?

| Kontrol | Kural | Önem |
|---|---|---|
| Şasi no | 17 karakter, harf ve rakam, I / O / Q yok | Ret riski |
| Garanti süresi | Arıza tarihi, başlangıç ile başlangıç + `--garanti-ay` (24) arasında | Ret riski |
| Km sınırı | `--garanti-km` verilmişse aşılmamalı | Ret riski |
| Başvuru süresi | Başvuru − onarım ≤ `--basvuru-gun` (30) | Ret riski |
| Kodlar | Arıza ve işçilik kodu dolu; kod listesi varsa listede olmalı | Düzeltilmeli / Ret riski |
| İşçilik süresi | Dolu olmalı; standart süreyi aşmamalı | Düzeltilmeli |
| Parça | Parça no varsa adet pozitif | Düzeltilmeli |
| Tarih sırası | Onarım arızadan, başvuru onarımdan önce olamaz | Düzeltilmeli |
| Mükerrer başvuru | Aynı başvuru numarası | Ret riski |
| Tekrarlayan onarım | Aynı şasi ve arıza kodu `--tekrar-gun` (30) içinde | Düzeltilmeli |
| Km tutarlılığı | Aynı şaside km sonraki tarihte daha düşük olamaz | Ret riski |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --basvurular basvurular.xlsx
python main.py --basvurular basvurular.xlsx --kodlar kodlar.xlsx --garanti-ay 36 --garanti-km 100000 --basvuru-gun 15
```

- Başvurular: Şasi No, Arıza Tarihi (zorunlu); Başvuru No, Garanti Başlangıç, Onarım Tarihi, Başvuru Tarihi, Km, Arıza
  Kodu, İşçilik Kodu, İşçilik Süresi, Parça No, Parça Adedi.
- Kodlar (isteğe bağlı): Kod Türü (Arıza / İşçilik), Kod, Standart Süre. Verilmezse kodların varlığı denetlenmez.

## Çıktı

`cikti/garanti_basvuru_kontrol.xlsx`:

- `Kontrol Sonucu`: başvuru bazında durum ve bulgular; düzeltme notu için boş sütun.
- `Bulgu Özeti`: bulgu türlerine göre adetler ve durum dağılımı.
- `Uyarılar`: her bulgu açıklamasıyla.

## Dikkat

- **Kurallar markaya göre değişir.** Garanti süresi, km sınırı, başvuru süresi, parça saklama ve fotoğraf şartları
  üreticinin garanti el kitabında yazar. Varsayılan değerler örnektir; parametreleri kendi markanıza göre verin.
- **Şasi kontrolü yalnız biçime bakar.** Numaranın o araca ait olduğunu veya kontrol hanesini doğrulamaz.
- **Teknik teşhis denetlenmez.** Arıza kodu ile değişen parçanın uyumu, kullanıcı hatası, bakım geçmişi ve
  kampanya (geri çağırma) kapsamı kod tarafından bilinmez.
- **Uzatılmış garanti ve iyi niyet (kulans) başvuruları** "garanti süresi dolmuş" olarak işaretlenir; onayınız varsa
  bulguyu dikkate almayın.
- **Tüketici mevzuatındaki yasal garanti hakları bu kontrolün konusu değildir.** Kod üreticiyle bayi arasındaki
  başvuru sürecine yardımcı olur.
- **Örnek veri:** Şasi numaraları, kodlar ve başvurular kurgusaldır.

## Testler

- Şasi biçimi, ay ekleme, her tekil kuralın sınır değerleri.
- Tekrarlayan onarım, km tutarsızlığı, mükerrer başvuru; kod listesi varken ve yokken davranış.
- Okuma (şasi temizleme, eksik alanlar), komut satırı hataları; örnek veri uçtan uca.

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
