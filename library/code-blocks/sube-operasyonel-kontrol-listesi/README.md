# Şube Operasyonel Kontrol Listesi · Kod Bloğu

> Şube Bankacılığı › Şube Operasyon Yöneticisi · Workers / Workless

Şubenin günlük, haftalık ve aylık periyodik kontrollerini izler; her sabah **bugün yapılacak ve gecikmiş
kontrolleri** sorumlusuyla listeler, paraf atılacak çıktıyı hazırlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Sonraki tarih, günlük | Son yapılma tarihinden sonraki ilk iş günü |
| Sonraki tarih, diğerleri | Haftalık: + 7 gün · Aylık: + 1 ay · Üç Aylık: + 3 ay · Altı Aylık: + 6 ay · Yıllık: + 12 ay |
| İş günü kaydırma | Sonraki tarih hafta sonuna ya da tatil listesindeki güne gelirse sonraki ilk iş gününe kayar |
| Durum | Hiç yapılmamış · Gecikti (durum tarihi sonraki tarihi geçmiş) · Bugün · Tamam |
| Uyum % | (Kontrol sayısı − geciken − hiç yapılmamış) ÷ kontrol sayısı × 100 |

**Örnek:** Üç aylık kontrol en son 10.07.2026'da yapılmış → sonraki tarih 10.10.2026 cumartesi → **12.10.2026
pazartesi**. Durum tarihi 12.10.2026 ise durum **Bugün**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --kontroller kontroller.xlsx
python main.py --kontroller kontroller.xlsx --tatiller tatiller.csv --tarih 30.10.2026
```

Zorunlu sütunlar: Kontrol, Periyot. Kontrolü yaptıkça "Son Yapılma Tarihi" sütununu güncellersiniz; sonuç
"Bulgu" ya da "Uygun Değil" yazılırsa rapor bunu ayrıca uyarır.

## Çıktı

`cikti/sube_operasyonel_kontrol_listesi.xlsx`:

- `Bugünün Listesi`: hiç yapılmamış, gecikmiş ve bugün yapılacak kontroller; tarih / paraf ve not sütunu boş gelir.
- `Tüm Kontroller`: her kontrolün son ve sonraki tarihi, durumu, son sonucu.
- `Periyot Özeti`, `Sorumlu Özeti`: duruma göre sayılar.
- `Uyarılar`.

## Dikkat

- **Kontrol listesi sizindir.** Hangi kontrolün hangi sıklıkla ve kim tarafından yapılacağını bankanızın iç
  düzenlemeleri, iç kontrol ve teftiş talimatları belirler. Örnek dosyadaki kontroller yalnız gösterim içindir; eksiksiz
  bir şube kontrol seti değildir.
- **Kod kontrolün yapıldığını doğrulamaz.** Yazdığınız tarihi esas alır; kasanın gerçekten sayıldığını, çift
  imzanın atıldığını ya da anahtarın doğru kişide olduğunu göstermez. Çıktı, imzalı tutanağın yerini tutmaz.
- **Tatil günlerini siz verirsiniz.** Tatil dosyası vermezseniz yalnız hafta sonları iş günü dışı sayılır; yarım
  günler tam iş günü kabul edilir.
- **Aylık ve daha seyrek kontrollerde tarih, son yapılma tarihine göre ilerler.** "Her ayın ilk iş günü" gibi
  takvime bağlı kontroller için son yapılma tarihini buna göre girin.
- **Kapsam dışı:** kasa ve efektif sayım farkı hesabı, kıymetli evrak seri numarası karşılaştırması (bunun için
  Kıymetli Evrak Sayım Kontrolü paketine bakın), görevler ayrılığı matrisi, bulgu aksiyon takibi.
- **Örnek veri:** Kontroller, sorumlular ve tarihler kurgusaldır.

## Testler

- Günlük kontrolde hafta sonu ve tatil atlama; haftalık ve aylık sonraki tarih; ay sonu taşması; iş gününe kaydırma.
- Durum sınırları (sonraki tarih dün / bugün / yarın), hiç yapılmamış kontrol, bulgulu sonuç, ileri tarih.
- Uyum yüzdesi, periyot ve sorumlu özeti, iş listesi sırası, periyot yazımları, okuma ve komut satırı.

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
