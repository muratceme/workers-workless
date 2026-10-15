# Periyodik Bakım Hatırlatma Listesi · Kod Bloğu

> Satış Sonrası Hizmetler (Servis) › Servis Danışmanı · Workers / Workless

Son bakım tarihi ve kilometresinden **bir sonraki bakımın ne zaman geleceğini** tahmin eder; bakımı geçmiş ya da yaklaşan araçları
listeler ve ileti izni olan müşteriler için **hatırlatma mesajı** hazırlar. Mesaj göndermez; internete ve DMS'e bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Günlük km | (Son bakım km − önceki bakım km) ÷ iki bakım arası gün; önceki bakım yoksa `--gunluk-km` (40) |
| Tahmini km (bugün) | Son bakım km + günlük km × son bakımdan bu yana geçen gün |
| Km'ye göre vade | Son bakım tarihi + ⌈bakım aralığı km ÷ günlük km⌉ gün |
| Süreye göre vade | Son bakım tarihi + bakım aralığı ay |
| Bakım vadesi | Hangisi önce gelirse (Vade Nedeni: Km / Süre) |
| Durum | Gecikmiş (vade geçti) · Yaklaşıyor (`--pencere` 30 gün içinde) · İleride |

Bakım aralığı satırda yazılı değilse `--aralik-km` (15.000) ve `--aralik-ay` (12) kullanılır. Mesaj şablonu `--mesaj` ile
değiştirilebilir; kullanılabilecek alanlar: `{musteri} {plaka} {tarih} {model} {firma}`. Vadesi gelmiş araçta `{tarih}` "geldi",
yaklaşanda "GG.AA.YYYY civarında geliyor" olur.

**Uyarılar:** önceki bakım km'si son bakımdan büyük ya da tarihler ters (Orta) · ileti izni var ama telefon yok (Orta) · bakım
vadesini 180 günden fazla geçmiş araç (başka servise gitmiş olabilir) · varsayılan bakım aralığı kullanılan araçlar · listede
olup ileti izni olmayan araçlar (Bilgi).

**Örnek:** 01.06.2025'te 6.000 km, 01.06.2026'da 30.000 km → günde 65,8 km; 15.000 km'lik aralık 228 günde dolar →
bakım vadesi **15.01.2027** (süreye göre vade 01.06.2027'den önce).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                            # örnek veri (14.10.2026)
python main.py --araclar araclar.xlsx --firma "Örnek Otomotiv Servis"
python main.py --araclar araclar.xlsx --pencere 45 --gunluk-km 50 --aralik-km 20000 --aralik-ay 12
python main.py --araclar araclar.xlsx --mesaj "Sayın {musteri}, {plaka} bakım zamanı {tarih}. {firma}"
```

Zorunlu sütunlar: Plaka ve Son Bakım Tarihi.

## Çıktı

- `cikti/bakim_hatirlatma.xlsx`: `Hatırlatma Listesi` (en gecikmiş üstte), `Tüm Araçlar`, `Mesajlar`, `Uyarılar`.
- `cikti/bakim_hatirlatma_mesajlar.csv`: Telefon; Mesaj (SMS / toplu mesaj aracına aktarmak için).

## Dikkat

- **İleti izni zorunlu kabul edilir.** Mesaj yalnız "İleti İzni" sütunu Evet olan müşterilere hazırlanır. 6563 sayılı Elektronik
  Ticaretin Düzenlenmesi Hakkında Kanun ve ilgili yönetmelik kapsamında ticari elektronik ileti için alıcının onayı ve İleti
  Yönetim Sistemi (İYS) kaydı gerekir. Bakım hatırlatmasının sizin durumunuzda ticari ileti sayılıp sayılmadığını ve İYS
  kayıtlarınızı hukuk / uyum biriminizle doğrulayın.
- **Kişisel veri içerir.** Müşteri adı, telefon ve plaka KVKK kapsamındadır; listeyi yalnız yetkili personelle paylaşın.
- **Bakım aralığını üretici belirler.** Model ve motora göre km / süre aralıkları, ağır kullanım koşulları ve servis göstergesi
  (değişken bakım aralığı) farklı olabilir; varsayılanlar yalnız örnektir.
- **Km tahmindir.** Kullanım alışkanlığı değişen araçlarda vade sapabilir; müşteri aranırken güncel km sorulmalıdır.
- **Örnek veri:** Müşteriler, plakalar ve telefonlar kurgusaldır.

## Testler

- Geçmiş bakımdan günlük km, km ve süre vadesi, hangisinin önce geldiği, tahmini km, durum ve sıralama, ay ekleme (ay sonu, artık yıl).
- Mesaj şablonu, izin ve telefon kontrolü, uyarılar, CSV çıktısı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |
| 6563 sayılı Elektronik Ticaretin Düzenlenmesi Hakkında Kanun | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6563.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

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
