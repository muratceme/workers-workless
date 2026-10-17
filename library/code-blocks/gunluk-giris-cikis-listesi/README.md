# Günlük Giriş-Çıkış Listesi · Kod Bloğu

> Ön Büro › Resepsiyonist · Gerisi Bende

Rezervasyon listesinden seçilen günün **gelecek (arrival), ayrılacak (departure) ve konaklamaya devam eden** misafirlerini
oda tipi, kişi sayısı, özel istek ve ödeme durumuyla listeler. Oda çakışmasını, oda atanmamış gelişleri ve çıkışta tahsil
edilecek bakiyeleri işaretler. İnternete ve PMS'e bağlanmaz; dışa aktarılmış rezervasyon listesini okur.

## Nasıl hesaplar?

| Liste / gösterge | Kural |
|---|---|
| Gelecekler | Giriş tarihi seçilen gün olan, iptal / no-show olmayan rezervasyonlar |
| Ayrılacaklar | Çıkış tarihi seçilen gün olanlar |
| Devam edenler | Giriş < gün < çıkış |
| Bakiye | Toplam tutar − ödenen |
| Ödeme durumu | Ödeme şeklinde *acente, cari, fatura, kurum, şirket, voucher* geçiyorsa "Acente / cari" (misafirden tahsil edilmez); değilse Ödendi / Kısmi ödeme / Ödenmedi |
| Gece sonu dolu oda | Gelecek + devam eden; aynı odadaki rezervasyonlar bir kez, odası atanmamış her rezervasyon bir oda sayılır |
| Doluluk % | Dolu oda ÷ `--oda-sayisi` × 100 |

Durum metinden tanınır: boş / Kesin / Garantili · Opsiyon · İptal / No-show · Konaklıyor / In house · Çıkış yaptı.

**Uyarılar:** aynı gece aynı odada iki rezervasyon (Yüksek) · oda sayısı aşıldı · oda atanmamış geliş · bugün gelecek
opsiyonlu rezervasyon · giriş tarihi geçtiği hâlde durumu "Kesin" kalan rezervasyon (Orta) · gelen misafirin odası aynı gün
boşalıyor (erken giriş riski) · çıkışta tahsil edilecek bakiye (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (15.10.2026)
python main.py --rezervasyonlar rezervasyonlar.xlsx --tarih 15.10.2026
python main.py --rezervasyonlar rezervasyonlar.xlsx --oda-sayisi 40     # tarih verilmezse bugün
```

Zorunlu sütunlar: Giriş Tarihi ve Çıkış Tarihi. Diğer sütunlar isteğe bağlıdır; boş bırakılanlar listede boş görünür.

## Çıktı

`cikti/giris_cikis_listesi.xlsx`: `Özet`, `Gelecekler` (odaya göre sıralı; odası atanmamışlar sonda, sarı), `Ayrılacaklar`,
`Devam Edenler`, `Uyarılar`.

## Dikkat

- **Bakiye yalnız konaklama tutarıdır.** Minibar, restoran ve diğer ekstralar PMS folyosundan kontrol edilmelidir.
- **Liste, dışa aktarıldığı andaki durumu gösterir.** Gün içinde gelen yeni rezervasyon ve iptaller için yeniden çalıştırın.
- **Kişisel veri içerir.** Misafir adı ve özel istekler (ör. sağlık, diyet) KVKK kapsamındadır; çıktıyı yalnız yetkili
  personelle paylaşın, basılı listeleri gün sonunda imha edin.
- **Örnek veri:** Misafir adları, odalar ve tutarlar kurgusaldır.

## Testler

- Gelecek / ayrılacak / devam eden ayrımı, dolu oda ve kişi sayısı, doluluk, ödeme durumu ve tahsil edilecek bakiye.
- Oda çakışması, atanmamış oda, opsiyon, giriş kaydı yok, aynı gün boşalan oda ve oda sayısı aşımı uyarıları.
- Durum tanıma, okuma hataları ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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
