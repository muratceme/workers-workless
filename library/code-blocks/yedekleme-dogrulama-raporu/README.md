# Yedekleme Doğrulama Raporu · Kod Bloğu

> Bilgi Teknolojileri › Sistem Uzmanı · Workers / Workless

Yedekleme yazılımınızdan dışa aktardığınız iş geçmişini planla karşılaştırır: **hangi yedek başarısız, hangisi
hiç çalışmamış, hangisi gecikmiş, hangisinin boyutu şüpheli derecede küçük.** Her sabah bakılacak tek tabloyu
üretir. İnternete bağlanmaz, sunuculara erişmez.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Geçerli yedek | Durumu "Başarılı" veya "Uyarı" olan koşu (uyarılı olanlar ayrıca işaretlenir) |
| Son yedek yaşı | Rapor anı − son geçerli yedeğin bitiş saati |
| Gecikmiş | Yaş > beklenen aralık + `--tolerans` (6 saat). Günlük 24, Haftalık 168, Aylık 744 saat |
| Başarı oranı | Son `--gun` (7) gündeki başarılı koşu ÷ tüm koşular |
| Boyut şüphesi | Son geçerli yedek, aynı türdeki önceki geçerli yedeklerin (en az 3) medyanından %50 veya daha fazla küçük |
| Süre aşımı | Son koşunun süresi plandaki "Azami Süre (dk)" değerinden uzun |
| Eksik | Planda olup geçmişte hiç koşusu olmayan iş |

İşin durumu: yüksek önemli bulgu varsa **Sorunlu**, orta önemli bulgu varsa **Dikkat**, yoksa **Tamam**.

**Örnek:** Son geçerli yedek 10.10 01:20, rapor anı 12.10 09:00 → yaş 55,7 saat. Günlük iş için sınır 24 + 6 = 30
saat → **gecikmiş**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                # örnek veri, rapor anı 12.10.2026 09:00
python main.py --gecmis gecmis.csv --plan plan.csv
python main.py --gecmis gecmis.csv --plan plan.csv --an "12.10.2026 09:00" --gun 14 --tolerans 12 --boyut-sapma 30
```

Geçmiş dosyasında zorunlu sütunlar: İş Adı, Başlangıç, Durum. Diğerleri: Sistem, Bitiş, Boyut (GB), Tür.
İngilizce başlıklar (Job Name, Start Time, End Time, Status, Size) ve durumlar (Success, Warning, Failed) da
tanınır. Plan dosyasında: İş Adı, Sıklık (Günlük / Haftalık / Aylık veya saat cinsinden sayı), isteğe bağlı
Sistem ve Azami Süre (dk). Plan verilmezse gecikme, eksik iş ve süre denetimi yapılmaz.

## Çıktı

`cikti/yedekleme_dogrulama.xlsx`:

- `İş Özeti`: her iş için durum, son geçerli yedek ve yaşı, son koşunun durumu, boyutu ve süresi, penceredeki
  başarı oranı. Sorunlu işler en üsttedir.
- `Sorunlu Koşular`: penceredeki başarısız ve uyarılı koşular, yeniden eskiye.
- `Uyarılar`: gecikmiş, eksik ve başarısız yedekler, boyut ve süre sapmaları, plan dışı işler.

## Dikkat

- **"Başarılı" geri yüklenebilir demek değildir.** Rapor yalnız yedekleme yazılımının kendi beyanını okur; yedeğin
  içeriğini açmaz. Düzenli geri yükleme testi yapın ve sonucunu ayrıca kaydedin.
- **Boyut denetimi kaba bir işarettir.** Sıkıştırma, tekilleştirme ve temizlik işleri boyutu meşru biçimde
  düşürebilir; tam ve artımlı yedekler yalnız kendi türleriyle karşılaştırılır (Tür sütunu doluysa).
- **Dışa aktarım biçimi üründen ürüne değişir.** Sütun adlarınız tanınmazsa başlıkları yukarıdaki adlara çevirin.
  Saat dilimi dönüştürülmez; geçmiş ve rapor anı aynı saat diliminde olmalıdır.
- **Eşikler tercihtir.** 6 saat tolerans ve %50 sapma genel başlangıç değerleridir; kendi kurtarma hedeflerinize
  (RPO) göre ayarlayın.
- **Kapsam dışı:** yedeğin başka konuma kopyalanması, saklama süresi, şifreleme ve depolama alanı denetimi.
- **Örnek veri:** Sistemler ve koşular kurgusaldır.

## Testler

- Yaş ve gecikme sınırı, hiç çalışmamış iş, geçerli yedeği olmayan iş, son koşu başarısız.
- Boyut medyanı ve sapma sınırı, sıfır boyut, süre aşımı, pencere sınırı, plan dışı iş.
- Tarih-saat ve durum okuma (Türkçe ve İngilizce), komut satırı.

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
