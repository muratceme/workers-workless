# Hasta Seyahat ve Tedavi Takvimi · Kod Bloğu

> Uluslararası Hasta Hizmetleri › Uluslararası Hasta Koordinatörü · Workers / Workless

Uçuş, transfer, konaklama, muayene, tetkik, operasyon, yatış ve kontrol kayıtlarından **hasta bazında takvim** çıkarır;
planı tutarlılık açısından kontrol eder, yaklaşan etkinlikleri listeler ve takvim uygulamalarına aktarılabilen bir
**.ics dosyası** üretir. İnternete bağlanmaz, mesaj göndermez.

## Nasıl kontrol eder?

| Kontrol | Kural |
|---|---|
| Randevu varıştan önce / dönüşten sonra | Muayene, tetkik, operasyon, yatış ya da kontrol, varış uçuşundan önce ya da dönüş uçuşundan sonra |
| Operasyon öncesi muayene yok | Operasyondan önce muayene kaydı yok |
| Aynı anda iki etkinlik | Saati yazılmış iki etkinlik aynı saatte başlıyor ya da süreleri çakışıyor |
| Konaklama boşluğu | Varış ile dönüş arasındaki bir gece ne konaklama ne yatışla karşılanıyor |
| Uçuş günü transfer yok | Varış ya da dönüş gününde transfer kaydı yok |
| Operasyon sonrası uçuş erken | Operasyon ile dönüş uçuşu arası `--ucus-sure` günden az (yalnız değer verilirse) |

Etkinlik türü metinden tanınır (ör. "Geliş uçuşu" → Varış uçuşu, "Ameliyat" → Operasyon, "Otel" → Konaklama). Saat
Başlangıç hücresinde ("10.10.2026 08:30") ya da ayrı Saat sütununda yazılabilir.

**Hatırlatmalar:** rapor tarihinden itibaren `--gun` (varsayılan 3) gün içindeki etkinlikler.

**Örnek:** Koroner anjiyografi 14.10.2026'da, dönüş uçuşu 17.10.2026'da → 3 gün; `--ucus-sure 5` verildiğinde "uçuş
uygunluğunu hekimle teyit edin" uyarısı. Aynı hastanın 15 ve 16 Ekim geceleri için konaklama ya da yatış kaydı yok.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                # örnek veri (09.10.2026, uçuş süresi 5 gün)
python main.py --etkinlikler hasta_etkinlikleri.xlsx --tarih 09.10.2026
python main.py --etkinlikler hasta_etkinlikleri.xlsx --gun 7 --ucus-sure 10
```

Zorunlu sütunlar: Hasta, Tür, Başlangıç. Konaklama ve yatışta Bitiş (çıkış günü) yazılmalıdır.

## Çıktı

- `cikti/hasta_takvimi.xlsx`: `Takvim` (hasta ve zamana göre), `Yaklaşanlar`, `Uyarılar`.
- `cikti/hasta_takvimi.ics`: tüm etkinlikler; Google Takvim, Outlook ve benzeri uygulamalara içe aktarılabilir. Saatler
  yerel saat olarak yazılır (saat dilimi belirtilmez).

## Dikkat

- **Tıbbi kararlar hekimindir.** Operasyon sonrası uçuş süresi, kontrol sayısı ve taburculuk hekim kararıdır; kod
  varsayılan süre kullanmaz, `--ucus-sure` yalnız verdiğiniz değeri kontrol eder.
- **Sağlık verisi içerir.** Takvim ve .ics dosyası özel nitelikli kişisel veri taşır; hasta adı yerine dosya numarası
  kullanın, dosyayı paylaşılan takvimlere aktarmadan önce erişimi sınırlayın. Yurt dışındaki hasta ya da aracıyla
  paylaşım KVKK'nın yurt dışı aktarım kurallarına tabidir.
- **Uçuş saatleri yerel saattir.** Farklı saat dilimlerinden gelen uçuşlarda varış saatini Türkiye saatiyle yazın.
- **Örnek veri:** Hasta numaraları, uçuşlar ve randevular kurgusaldır.

## Testler

- Varış / dönüş dışı randevu, operasyon öncesi muayene, çakışma, konaklama boşluğu, transfer ve uçuş süresi kontrolleri.
- Tür tanıma, tarih-saat okuma, yaklaşan etkinlikler, .ics üretimi, okuma ve komut satırı.

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
