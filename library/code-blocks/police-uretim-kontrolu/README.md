# Poliçe Üretim Kontrolü · Kod Bloğu

> Teknik (Risk Kabul / Underwriting) › Teknik Uzman Yardımcısı · Workers / Workless

Üretilen poliçeleri **onaylı tekliflerle** teminat satırı düzeyinde karşılaştırır; teminat, bedel, prim, muafiyet,
tarih ve özel şart uyumsuzluklarını listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

Poliçe, `Teklif No` ile teklife; teminatlar adlarıyla birbirine eşlenir (büyük-küçük harf ve Türkçe karakter ayrımı
yapılmaz). Aynı numaralı satırlar tek belge sayılır; her satır bir teminattır.

| Kontrol | Kural | Düzey |
|---|---|---|
| Teklif bulunamadı | Poliçedeki teklif numarası teklif listesinde yok | Yüksek |
| Aynı teklife birden çok poliçe | Aynı teklif numarası birden çok poliçede | Yüksek |
| Teminat poliçede yok / teklifte yok | Teminat yalnız bir tarafta | Yüksek |
| Bedel farklı, prim farklı | Fark `--tolerans` TL'den (varsayılan 1) büyük veya bir taraf boş | Yüksek |
| Başlangıç / bitiş tarihi farklı | İki tarafta da dolu ve farklı | Yüksek |
| Özel şart poliçede yok | Teklifteki özel şart / kloz poliçede yazmıyor | Yüksek |
| Muafiyet farklı | Muafiyet metni farklı | Orta |
| Sigortalı adı farklı | İki tarafta da dolu ve farklı | Orta |
| Özel şart teklifte yok | Poliçeye onayda olmayan özel şart eklenmiş | Orta |

- Özel şartlar `|`, `;` veya satır sonu ile ayrılır; sıra önemli değildir, metinler birebir (harf duyarsız) eşleşmelidir.
- Hiç farkı olmayan poliçe "Uyumlu" sayılır. Poliçesi bulunmayan teklifler ayrı sayfada listelenir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --teklifler teklifler.xlsx --policeler policeler.xlsx --tolerans 5
```

| Dosya | Sütunlar |
|---|---|
| Teklifler | Teklif No, Sigortalı, Teminat, Bedel, Prim, Muafiyet, Başlangıç, Bitiş, Özel Şartlar |
| Poliçeler | Poliçe No, Teklif No, Sigortalı, Teminat, Bedel, Prim, Muafiyet, Başlangıç, Bitiş, Özel Şartlar |

Sigortalı, tarih ve özel şart bilgisi her satırda tekrar edebilir veya yalnız ilk satırda olabilir.

## Çıktı

`cikti/police_uretim_kontrol.xlsx`:

- `Poliçe Özeti`: poliçe bazında sonuç, fark sayısı, teklif ve poliçe primi; boş "Kontrol Eden" sütunu.
- `Farklar`: her fark için teklifteki ve poliçedeki değer; boş "Düzeltildi" sütunu.
- `Üretilmemiş Teklifler`: poliçesi bulunmayan onaylı teklifler.
- `Uyarılar`: okunamayan tutar ve tarihler, tekrar eden teminatlar.

## Dikkat

- **Tablo karşılaştırmasıdır.** Kod poliçe metnini (PDF) okumaz; iki tabloya aktarılan alanları karşılaştırır.
  Tabloya girilmeyen bir fark bulunamaz.
- **Özel şartlar metin olarak eşleşir.** Aynı kloz iki tarafta farklı yazılmışsa (kısaltma, farklı ifade) fark olarak
  görünür; kod klozun içeriğini yorumlamaz.
- **Prim karşılaştırması** girdiğiniz tutar türüne göredir (net / brüt). İki dosyada aynı prim tanımını kullanın;
  vergi ve fon hesaplanmaz.
- **Zeyil ve yenilemeler:** Aynı teklife bağlı birden çok poliçe mükerrer üretim olarak işaretlenir; zeyilleri ayrı
  numarayla veya ayrı dosyada izleyin.
- **Örnek veri:** Sigortalılar, teminatlar, bedeller ve primler kurgusaldır.

## Testler

- Uyumlu poliçe, tolerans sınırı, boş prim.
- Eksik / fazla teminat, muafiyet, özel şartlar (sıra ve harf duyarsız).
- Sigortalı ve tarih farkları; teklifi olmayan poliçe, mükerrer üretim, üretilmemiş teklif.
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
