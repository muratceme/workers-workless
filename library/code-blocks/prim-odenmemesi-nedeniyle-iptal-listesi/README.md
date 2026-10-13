# Prim Ödenmemesi Nedeniyle İptal Listesi · Kod Bloğu

> Tahsilat Yönetimi › Tahsilat Uzmanı · Workers / Workless

Vadesi geçtiği hâlde ödenmeyen prim taksitlerini **süre kurallarına göre sınıflar**: ihtar gönderilecek poliçeleri,
ihtar süresi işleyenleri ve süresi dolduğu hâlde ödenmeyen **iptal adaylarını** listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

Gecikmiş taksit: kalan tutarı (tutar − ödenen) sıfırdan büyük ve vadesi durum tarihinden önce olan taksit.

| Taksit | Durum | Kural |
|---|---|---|
| İlk taksit (peşinat) | Peşinat ödenmedi | Vade + `--cayma-ay` (3) ay henüz dolmadı |
| İlk taksit (peşinat) | Peşinat ödenmedi (süre doldu) | Durum tarihi > vade + 3 ay |
| İzleyen taksitler | İhtar gönderilecek | İhtar tebliğ tarihi boş |
| İzleyen taksitler | İhtar süresi işliyor | Durum tarihi ≤ ihtar tebliğ tarihi + `--ihtar-sure` (10) gün |
| İzleyen taksitler | İptal adayı | Durum tarihi > ihtar tebliğ tarihi + 10 gün ve taksit hâlâ ödenmemiş |

Poliçenin durumu, gecikmiş taksitlerinin en ağır olanıdır (sıra: İptal adayı, Peşinat süre doldu, Peşinat
ödenmedi, İhtar süresi işliyor, İhtar gönderilecek). `Taksit No` sütunu boşsa taksitler vade sırasına göre
numaralanır.

**Örnek:** 3. taksitin vadesi 10.08.2026, ihtar 25.09.2026'da tebliğ edilmiş → süre **05.10.2026**'da doldu.
Durum tarihi 12.10.2026 ve taksit ödenmemiş → poliçe **İptal adayı**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --taksitler taksitler.xlsx
python main.py --taksitler taksitler.xlsx --ihtar-sure 10 --cayma-ay 3 --gun 30 --tarih 31.10.2026
```

Zorunlu sütunlar: Poliçe No, Vade, Tutar. Her satır bir taksittir.

## Çıktı

`cikti/iptal_adayi_policeler.xlsx`:

- `Poliçeler`: gecikmiş taksiti olan poliçeler, en ağır durumdan başlayarak; gecikmiş tutar, süre sonu, taksit ayrıntısı.
- `Özet`: durum ve branş bazında poliçe sayısı ve gecikmiş tutar.
- `Taksitler`: bütün taksitlerin durumu.
- `Uyarılar`.

## Dikkat

- **Süreler çalışma varsayımıdır; kod iptal kararı vermez.** Varsayılanlar, Türk Ticaret Kanunu'nun prim borcunda
  temerrüde ilişkin genel düzenlemesindeki yaklaşımı izler: ilk taksit ödenmezse vadeden başlayan üç aylık cayma
  süresi; izleyen taksitlerde ihtarla verilen on günlük süre. Uygulamadan önce güncel kanun metnini, ilgili branşın
  genel şartlarını ve poliçe özel şartlarını hukuk biriminizle doğrulayın.
- **Zorunlu sigortalarda ve bazı branşlarda kurallar farklıdır.** Zorunlu sorumluluk sigortaları, hayat ve sağlık
  sigortaları gibi ürünlerde prim ödenmemesinin sonuçları ayrı düzenlenmiş olabilir; bu poliçeleri aynı
  varsayımlarla değerlendirmeyin ya da süreleri parametrelerle değiştirerek ayrı çalıştırın.
- **İhtarın geçerliliğini kod bilemez.** İhtarın hangi yolla gönderildiği, içeriği ve gerçekten tebliğ edilip
  edilmediği kontrol edilmez; dosyaya yazdığınız tebliğ tarihi esas alınır. Tebliğ tarihini değil gönderim tarihini
  yazarsanız süre erken dolmuş görünür.
- **"Süre doldu" hukuki sonucu göstermez.** İlk taksit ödenmeyen poliçede sorumluluğun başlayıp başlamadığı,
  cayma hakkının kullanılıp kullanılmadığı ve primin talep edilip edilmediği ayrıca değerlendirilmelidir.
- **Kapsam dışı:** iptal zeylinin düzenlenmesi, gün esasına göre prim iadesi ya da kazanılmış prim hesabı, rehin
  alacaklısı ve dain-i mürtehine bildirim, gecikme faizi, kredi kartı ve otomatik ödeme iadeleri.
- **Örnek veri:** Poliçeler, sigortalılar ve tutarlar kurgusaldır.

## Testler

- İhtar süresi sınırı (sürenin son günü işliyor, ertesi gün iptal adayı), cayma süresi sınırı ve ay sonu taşması.
- Vadesi bugün olan taksit, kısmi ödeme, numarasız taksitlerin vadeye göre sıralanması, poliçenin en ağır durumu.
- Fazla ödeme, ihtar tarihinin vadeden önce olması, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
