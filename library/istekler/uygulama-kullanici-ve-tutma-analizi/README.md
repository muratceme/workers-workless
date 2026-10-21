# Uygulama Kullanıcı ve Tutma Analizi · Kod Bloğu

> Sizden Gelenler › Abonelikli mobil uygulama · Gerisi Bende

Firebase'in (Google Analytics) BigQuery olay dökümünden **günlük aktif kullanıcı**, **yeni kullanıcı**, **en çok
kullanılan özellikler** ve **D1 / D7 / D30 tutma oranlarını** çıkarır. İnternete bağlanmaz; Firebase hesabınıza
erişmez, yalnız BigQuery'den aldığınız dosyayı okur.

## Tanımlar

| Gösterge | Nasıl hesaplanır |
|---|---|
| Gün | `event_date` (uygulamanın saat dilimi); yoksa `event_timestamp`'in UTC günü |
| Kullanıcı | `user_pseudo_id` (uygulama kurulumu); yoksa `user_id` |
| Aktif kullanıcı (DAU) | O gün en az bir olayı olan kullanıcı. Arka plan olayları (`notification_receive`, `notification_dismiss`, `app_remove`) aktiflik saymaz (`--pasif-olaylar`) |
| Kohort | İlk açılış günü: `first_open` olayının günü; yoksa `user_first_touch_timestamp`; o da yoksa dökümde ilk görüldüğü gün (uyarı verilir) |
| DN tutma | Kohorttaki kullanıcılardan kohort gününden **tam N gün sonra** aktif olanların payı. N. gün dökümün dışındaysa boş kalır |
| Ortalama tutma | Verisi tamamlanmış kohortların kullanıcı sayısıyla ağırlıklı ortalaması |
| WAU / MAU | Son günde biten 7 / 30 günde farklı aktif kullanıcı · yapışkanlık = DAU ÷ MAU |
| Özellik | Otomatik toplanan olaylar (`session_start`, `user_engagement`, `screen_view`… `--otomatik`) dışındaki olaylar, yani uygulamanın kendi kaydettiği olaylar |

Bu tanımlar Google Analytics panosundaki "aktif kullanıcı" tanımıyla birebir aynı değildir; rakamlar panodan biraz
farklı çıkabilir. Tanımlar dosyada açıkça yazılı olduğu için dönemler arası karşılaştırma tutarlıdır.

## Dökümü nasıl alırsınız?

Google Analytics → BigQuery aktarımı açıkken olaylar günlük `events_YYYYMMDD` tablolarında durur. Ham döküm çok büyük
olabileceği için BigQuery'de **gün × kullanıcı × olay** özetini alın (son 60 gün; D30 için en az 60 gün gerekir):

```sql
SELECT event_date, user_pseudo_id, event_name, platform,
       app_info.version AS app_version, COUNT(*) AS olay_sayisi
FROM `PROJE_KIMLIGI.ANALYTICS_VERI_KUMESI.events_*`
WHERE _TABLE_SUFFIX BETWEEN FORMAT_DATE('%Y%m%d', DATE_SUB(CURRENT_DATE(), INTERVAL 60 DAY))
                        AND FORMAT_DATE('%Y%m%d', CURRENT_DATE())
GROUP BY 1, 2, 3, 4, 5
```

Ham döküm de olur: her satır bir olay sayılır (`olay_sayisi` sütunu yoksa 1).

## Kullanım

```bash
python main.py                                          # örnek dökümle dener (kurgusal, 17.08–30.09.2026)
python main.py --girdi firebase_olay_ozeti.csv
python main.py --girdi dokumler/ --tarih 30.09.2026 --otomatik "session_start,user_engagement,screen_view"
```

Terminal kullanmıyorsanız klasördeki `BASLAT.bat` (Windows) ya da `BASLAT.command` (macOS) dosyasına çift tıklayın.

## Çıktı

- `cikti/kullanici_ve_tutma.xlsx`: Özet, Günlük, Tutma (kohort tablosu), Özellikler (7 günde %30'dan fazla düşenler
  kırmızı), Platform ve Sürüm, Uyarılar.
- `cikti/uygulama-kullanici-ve-tutma-analizi.ozet.json`: **Uygulama Yönetim Özeti** paketinin okuduğu günlük
  sayılar, ortalama tutma ve en çok kullanılan özellikler.

## Kaynaklar ve dayanaklar

Kontrol tarihi: 21.10.2026. Alan adları ve otomatik olay adları aşağıdaki belgelerden alınmıştır; kendi
veri kümenizin şemasıyla karşılaştırın.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| Google Analytics for Firebase: BigQuery Export şeması (event_date, event_timestamp, user_pseudo_id, user_first_touch_timestamp, app_info.version) | [support.google.com](https://support.google.com/firebase/answer/7029846) | Alan adları ve birimleri belgeyle karşılaştırıldı |
| Firebase Analytics ayrılmış olay adları (first_open, session_start, user_engagement) | [firebase.google.com](https://firebase.google.com/docs/reference/swift/firebaseanalytics/api/reference/Classes/Analytics) | Olay adları belgeyle karşılaştırıldı; otomatik olay listesinin geri kalanı düzenlenebilir varsayılandır |

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
