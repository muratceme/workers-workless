# Uygulama Çökme Önceliklendirme · Kod Bloğu

> Sizden Gelenler › Abonelikli mobil uygulama · Gerisi Bende

Firebase Crashlytics'in BigQuery dökümünden son günlerin çökmelerini (FATAL), ANR'leri ve çökmeyen hatalarını
**P1 / P2 / P3** diye önceliklendirir: kaç kurulumu etkilediğine, en yeni sürümle mi geldiğine ve son gün sıçrama olup
olmadığına bakar. Sürüm bazında çökmesiz kullanıcı oranını çıkarır ve ekibe gönderilecek kısa **bildirim metnini**
yazar. İnternete bağlanmaz, bildirimi kendisi göndermez.

## Öncelik kuralları

| Öncelik | Kural |
|---|---|
| P1 | Çökme ya da ANR **ve** (etkilenen kurulum ≥ `--esik` (10) **ya da** en yeni sürümde ortaya çıkan ve ≥ 3 kurulum **ya da** sıçrama) |
| P2 | Diğer çökme / ANR · çökmeyen hatada sıçrama ya da ≥ 5 × `--esik` kurulum |
| P3 | Diğerleri |

- **Pencere:** dökümdeki son günde biten `--gun` (7) gün; `--tarih GG.AA.YYYY` ile değiştirilebilir.
- **Etkilenen kurulum:** farklı `installation_uuid` sayısı. Crashlytics kurulum bazında sayar; bir kişi uygulamayı
  silip yeniden kurarsa iki kurulum olur.
- **Yeni sorun:** ilk kez penceredeki en yeni `display_version`'da görülen sorun (2.10 > 2.9 sayısal sıralanır).
- **Sıçrama:** son gündeki olay ≥ önceki günlerin günlük ortalamasının 2 katı ve ≥ 5.
- **Çökmesiz kullanıcı (yaklaşık):** 1 − çöken kurulum ÷ sürümün aktif kullanıcısı. Yalnız `--kullanicilar` verilirse
  hesaplanır; kurulum ≈ kullanıcı varsayımına dayanır, Crashlytics panosundaki rakamla birebir aynı olmayabilir.
- Aynı `event_id` iki kez gelirse bir kez sayılır.

## Dökümü nasıl alırsınız?

Crashlytics → BigQuery aktarımı açıkken veriler `firebase_crashlytics` veri kümesinde, uygulama başına bir tabloda
durur (tablo adı uygulama kimliğinden türetilir, sonuna `_ANDROID` ya da `_IOS` eklenir). BigQuery'de aşağıdaki
sorguyu çalıştırıp sonucu CSV olarak kaydedin:

```sql
SELECT
  event_timestamp, event_id, issue_id, error_type, platform,
  application.display_version AS display_version,
  device.model AS device_model,
  operating_system.display_version AS os_version,
  installation_uuid,
  blame_frame.symbol AS title          -- isteğe bağlı: sorun başlığı olarak suçlanan kare
FROM `PROJE_KIMLIGI.firebase_crashlytics.UYGULAMA_TABLOSU_ANDROID`
WHERE event_timestamp >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 14 DAY)
```

Alan adları Firebase'in Crashlytics BigQuery şeması belgesinden alındı. `is_fatal` alanı kullanımdan kalkmıştır
(`error_type` kullanın); eski dökümlerde `is_fatal` varsa paket onu da okur. Başlık alanı yoksa sorunlar `issue_id`
ile gösterilir. İki platformunuz varsa iki tabloyu `UNION ALL` ile birleştirin ya da iki CSV'yi birlikte verin.

## Kullanım

```bash
python main.py                                                 # örnek dökümle dener (kurgusal, Eylül 2026)
python main.py --girdi crashlytics_dokumu.csv --kullanicilar surum_kullanicilari.csv
python main.py --girdi dokumler/ --gun 14 --esik 25 --tarih 30.09.2026
```

Terminal kullanmıyorsanız klasördeki `BASLAT.bat` (Windows) ya da `BASLAT.command` (macOS) dosyasına çift tıklayın.

## Çıktı

- `cikti/cokme_raporu.xlsx`: Kritik (P1 / P2), Sorunlar (hepsi), Sürümler (en yeni sürüm sarı), Günlük, Kurallar, Uyarılar.
- `cikti/cokme_bildirimi.txt`: Slack, Teams ya da e-postaya yapıştırılacak kısa özet (son gün sayıları ve P1 sorunlar).
- `cikti/uygulama-cokme-onceliklendirme.ozet.json`: **Uygulama Yönetim Özeti** paketinin okuduğu günlük sayılar ve P1 listesi.

## Kaynaklar ve dayanaklar

Kontrol tarihi: 21.10.2026. Alan adları aşağıdaki belgeden alınmıştır; Firebase şemayı değiştirebilir, kendi
tablonuzun şemasıyla karşılaştırın.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| Firebase Crashlytics: BigQuery veri kümesi şeması (event_timestamp, issue_id, error_type, installation_uuid, application.display_version, blame_frame) | [firebase.google.com](https://firebase.google.com/docs/crashlytics/bigquery-dataset-schema) | Alan adları belgeyle karşılaştırıldı |

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
