# SGK Teşvik Uygunluk Ön Kontrolü

> Bordro ve SGK İşlemleri › Bordro Elemanı · Workers / Workless

Çalışan ve işyeri bilgilerine göre yaygın **SGK prim teşviklerinin uygunluk koşullarını** ön kontrolden geçirir. Her
çalışan ve her teşvik için **Uygun / Uygun değil / Bilgi eksik** sonucunu nedenleriyle verir. 4447 geçici 10 teşvikinde
süreyi ve bitiş tarihini de hesaplar. Yapay zekâ kullanmaz; internete bağlanmaz.

## Kontrol edilen teşvikler

| Teşvik | Dayanak | Kontrol edilen koşullar |
|---|---|---|
| MYÖ işveren hissesinde puan indirimi | 5510 md. 81/1-(ı) | Özel sektör, bildirge ve prim yasal sürede, borç yok ya da yapılandırılmış, kayıt dışı / sahte bildirim tespiti yok. **2026: imalat 5 puan, imalat dışı 2 puan** (7566 sayılı Kanun). |
| Genç, kadın ve mesleki belgeli istihdam | 4447 geçici md. 10 | İşe giriş 01.03.2011-31.12.2026; 18 yaşını doldurmuş; işe girişten önceki 6 ayda sigortalı olarak kayıtlı değil; önceki 6 ayın ortalama sigortalı sayısına ilave; ortak işyeri şartları. |
| Engelli sigortalı istihdamı | 4857 md. 30 | Engelli sigortalı; bildirge yasal sürede. |
| İşsizlik ödeneği alanların istihdamı | 4447 md. 50 | İşe giriş 01.10.2009 ve sonrası; işe girişte işsizlik ödeneği hakkı; ödenekten önceki son işyeri bu işyeri değil; ortalamaya ilave; bildirge yasal sürede. |

**4447 geçici 10 süreleri** (SGK teşvik broşürü):

| Sigortalı | Mesleki yeterlilik belgesi | Mesleki ve teknik eğitim / İŞKUR kursu | Belgesiz |
|---|---|---|---|
| Kadın (18+) ya da 29 yaşını doldurmamış erkek | 48 ay | 36 ay | 24 ay |
| 29 yaşını doldurmuş erkek | 24 ay | 24 ay | İŞKUR kayıtlıysa 6 ay, değilse yok |

İŞKUR'a kayıtlı işsizler arasından işe alınanlarda (6 aylık satır hariç) **6 ay eklenir**. 4/1-a kapsamında çalışırken
belge alanlara uygulanan 12 aylık süre bu blokta yoktur.

## Kullanım

```bash
python main.py                                                         # örnek (kurgusal) 7 çalışan
python main.py --calisanlar yeni_girisler.xlsx --isyeri isyeri.txt --parametre tesvik_parametreleri.json
```

`isyeri.txt` örneği:

```
Kontrol Tarihi: 15.10.2026
Sektör: İmalat
Özel Sektör: Evet
Bildirgeler Yasal Sürede: Evet
Prim Borcu: Yapılandırılmış
Kayıt Dışı Tespiti: Hayır
```

## Çıktı

- `cikti/tesvik_uygunluk.xlsx`: Uygunluk (yeşil: uygun, kırmızı: uygun değil, sarı: bilgi eksik), Özet.

## Parametreler

Teşviklerin tarih aralıkları, puanları ve süreleri `ornek_veri/tesvik_parametreleri.json` dosyasındadır. Mevzuat
değiştiğinde (ör. son işe alım tarihinin uzatılması, puan değişikliği) bu dosyayı güncelleyin ve `son_teyit` alanına
kontrol tarihinizi yazın. Rapor bu tarihi gösterir.

## Dikkat

- **Ön kontroldür.** Kesin uygunluk SGK sistemindeki kayıtlara (ortalama sigortalı sayısı, önceki sigortalılık, İŞKUR
  kaydı, borç durumu) göre belirlenir. Teşvik tanımlamasından ve bildirgeden önce e-SGK teşvik sorgulama ekranlarıyla
  doğrulayın; gerekirse SMMM'nize danışın.
- **Birlikte uygulama.** Aynı sigortalı ve aynı prim günü için birden fazla teşvikten birlikte yararlanılıp
  yararlanılamayacağı her teşvikin kendi mevzuatına göre değerlendirilir; bu blok yalnız tek tek uygunluğu gösterir.
- **Kapsam dışı:** Yatırım teşvik belgesi (5510 ek 2), Ar-Ge / tasarım (5746), kültür (5225), yurt dışı (81/1-i),
  asgari ücret desteği ve bölgesel teşvikler bu blokta yoktur. Geriye dönük teşvik başvurularında süre sınırları ayrıca
  değerlendirilir.
- **Haksız teşvik.** Şartları sağlamadan alınan teşvik gecikme cezası ve zammıyla geri alınır; kayıt dışı istihdam
  tespitinde teşviklerden belirli süre yararlanılamaz.
- **Örnek veri:** Çalışanlar ve işyeri kurgusaldır.

## Testler

Testler; yaş (29 yaşını doldurma sınırı), ay ekleme, 4447 geçici 10'un tüm süre seçeneklerini (24, 30, 36, 48, 54, 6 ay),
tarih aralığı, 18 yaş, önceki 6 ay sigortalılık, ilave istihdam, süre dolması ve işyeri engellerini; 81/ı puanlarını,
engelli ve işsizlik ödeneği teşviklerini doğrular.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |
| 5510 sayılı Sosyal Sigortalar ve Genel Sağlık Sigortası Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.5510.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |
| 4447 sayılı İşsizlik Sigortası Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4447.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

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
