# Mükellef Beyanname Takvimi · Kod Bloğu

> Beyanname ve Vergi › Serbest Muhasebeci Mali Müşavir (SMMM) · Workers / Workless

Mükelleflerin yükümlülüklerinden **seçilen ayın beyanname ve ödeme takvimini** çıkarır; verilenleri, yaklaşanları ve
gecikenleri mükellef bazında izler. İnternete bağlanmaz.

## Nasıl çalışır?

Beyanname türlerini ve günlerini **siz tanımlarsınız**; kod bu tanımı mükellef listesine uygular.

| Sütun | Anlamı | Örnek |
|---|---|---|
| Periyot | Aylık, 3 Aylık veya Yıllık | Aylık |
| Dönem Sonu Ayları | Boşsa periyottan gelir (3 Aylık: 3, 6, 9, 12; Yıllık: 12) | 3,6,9,12 |
| Kaç Ay Sonra | Dönem bittikten kaç ay sonra verildiği (boşsa 1) | 2 |
| Beyan Günü / Ödeme Günü | Ayın günü veya `son` | 28 |

- Takvim ayından "kaç ay sonra" kadar geri gidilen ay, türün dönem sonu aylarından biriyse beyanname o ay verilir.
- Son gün cumartesi, pazar veya tatil dosyanızdaki bir güne rastlarsa izleyen ilk iş gününe kayar.
- Durum: **Verildi**, **Geç verildi**, **Gecikmiş**, **Yaklaşıyor** (`--yaklasan` gün veya daha az), **Bekliyor**.
  Ödeme günü tanımlıysa ödeme için de aynı durumlar izlenir.

**Örnek:** "SGK Prim Ödemesi; Aylık; kaç ay sonra 1; ödeme günü son" → Ekim 2026 takviminde Eylül dönemi; ayın son günü
31.10.2026 cumartesi olduğundan son gün 02.11.2026 pazartesi. (Form Ba-Bs, 565 Sıra No'lu VUK Genel Tebliği ile Eylül 2024
döneminden itibaren kaldırıldığı için örnekten çıkarıldı.)

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, Ekim 2026, bugün 26.10.2026
python main.py --yukumlulukler yukumlulukler.xlsx --turler beyanname_turleri.xlsx --ay 2026-10
python main.py --yukumlulukler yukumlulukler.xlsx --turler beyanname_turleri.xlsx --durum durum.xlsx --tatiller tatiller.xlsx --ay 2026-10 --yaklasan 5
```

- Yükümlülükler: Mükellef, Beyanname, Sorumlu.
- Durum (isteğe bağlı): Mükellef, Beyanname, Dönem (`2026-09`, `09.2026`, `Eylül 2026`), Verildi, Ödendi (tarih).
- Tatiller (isteğe bağlı): Tarih, Açıklama.

## Çıktı

`cikti/beyanname_takvimi.xlsx`:

- `Takvim`: son gün sırasıyla beyannameler; mükellef sayısı, verilen, kalan, geciken.
- `Mükellef Bazında`: her mükellefin o ayki beyannameleri, durumu ve ödeme durumu.
- `Uyarılar`: gecikmiş ve yaklaşan beyannameler, gecikmiş ödemeler, tanımsız türler, okunamayan satırlar.

## Dikkat

- **Kod vergi takvimini bilmez.** Örnek dosyadaki günler örnektir. Beyan ve ödeme süreleri mevzuatla değişir; GİB sık sık
  süre uzatır. Her ay GİB vergi takviminden doğrulayıp türler dosyanızı güncel tutun.
- **Resmi tatilleri siz girersiniz.** Kod yalnız hafta sonunu bilir; dini bayramlar, yarım günler ve idari izinler
  için tatiller dosyasını doldurun. Yarım gün tatillerde sürenin nasıl işlediğini ayrıca kontrol edin.
- **Özel durumlar kapsam dışıdır.** Özel hesap dönemi, tasfiye, işe başlama ve bırakma, mücbir sebep hâli için
  tarihleri elle yönetin.
- **"Verildi" kaydı sizin beyanınızdır.** Kod tahakkuk fişini veya ödemeyi doğrulamaz.
- **Örnek veri:** Mükellefler ve personel kurgusaldır.

## Testler

- Aylık, üç aylık ve yıllık türlerin doğru aya düşmesi; yıl geçişi; ayın son günü.
- Hafta sonu ve tatil kaydırması; verildi, geç verildi, gecikmiş, yaklaşıyor durumları; ödeme durumu.
- Tanımsız tür, okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 213 sayılı Vergi Usul Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.4.213.pdf) | Resmî metinle karşılaştırılan maddeler: 274, 280, 315, mükerrer 315, 320, 323 |
| Form Ba-Bs'nin Eylül 2024 döneminden itibaren kaldırılması — 565 Sıra No'lu VUK Genel Tebliği (RG 25.09.2024 / 32673) | [Andersen](https://tr.andersen.com/tr/images/pdf/and_41-ba-bs-formlari-kaldirildi.pdf), [İSMMMO](https://ismmmo.org.tr/dosya/5020/Mevzuat-Dosya/25092024-vuk565-sn-gn-tb-babs-kaldirildi.pdf), [Alomaliye](https://www.alomaliye.com/2024/09/25/vergi-usul-kanunu-genel-tebligi-sira-no-565/amp/) | Teyit edildi (ikincil kaynaklar) |

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
