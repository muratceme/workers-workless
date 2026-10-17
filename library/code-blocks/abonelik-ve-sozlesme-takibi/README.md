# Abonelik ve Sözleşme Takibi · Kod Bloğu

> İdari İşler › İdari İşler Sorumlusu · Gerisi Bende

Kira, enerji, su, internet, bakım, hizmet alımı ve kiralama sözleşmelerinin **tutar, bitiş, fesih ihbarı ve fiyat artış
tarihlerini** izler. Fesih ya da yeniden pazarlık için son bildirim gününü kaçırmamak üzere yaklaşanları listeler; kategori
bazında aylık ve yıllık maliyeti çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| İhbar son günü | Bitiş − fesih ihbar süresi (gün); yalnız otomatik yenilenen sözleşmelerde (bu güne kadar bildirim yapılmazsa sözleşme yenilenir) |
| Durum | Süresi doldu · İhbar yaklaşıyor (≤ `--ihbar` 30 gün) · Bitiş yaklaşıyor (≤ `--bitis` 60 gün) · İhbar süresi geçti · Devam ediyor · Süresiz (bitiş yok) |
| Aylık karşılık | Tutar × periyot katsayısı (yıllık ÷ 12, 3 aylık ÷ 3, 6 aylık ÷ 6) |
| Maliyet | Kategori ve para birimi bazında aylık ve yıllık toplam (süresi dolmuş, yenilenmeyen sözleşmeler hariç) |
| Artış tarihi | Artış esası yazılı sözleşmelerde başlangıcın bir sonraki yıl dönümü |

**Uyarılar:** süresi dolan ve otomatik yenilenmeyen sözleşme · okunamayan satır (Yüksek) · ihbar son günü yaklaşan · bitişi yaklaşan ·
süresi dolup otomatik yenilenen (bitiş tarihi güncellenmeli) (Orta) · ihbar süresi geçmiş · yaklaşan fiyat artış dönemi (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (14.10.2026)
python main.py --sozlesmeler sozlesmeler.xlsx
python main.py --sozlesmeler sozlesmeler.xlsx --ihbar 45 --bitis 90
```

## Çıktı

`cikti/sozlesme_takibi.xlsx`: `Sözleşmeler` (en acil üstte), `Maliyet`, `Uyarılar`.

## Dikkat

- **İhbar ve yenileme koşulları sözleşmeye bağlıdır.** Süreyi ve bildirim şeklini (noter, KEP, yazılı) sözleşme metninden alın.
- **Kira sözleşmeleri özel kurallara tabidir.** Türk Borçlar Kanunu'nun konut ve çatılı işyeri kiralarına ilişkin hükümleri (süre
  sonunda uzama, kiracının bildirim süresi, kira artışının sınırı) sözleşme maddesinden önce gelebilir. Artış oranını ve fesih
  hakkını güncel mevzuattan ve hukuk danışmanınızdan doğrulayın.
- **Tutarlar aynı esasta olmalı** (KDV dahil ya da hariç). Döviz tutarları çevrilmez, para birimi bazında ayrı toplanır.
- **Örnek veri:** Tedarikçiler, tutarlar ve tarihler kurgusaldır.

## Testler

- Durum sınıflandırması (otomatik yenilenen / yenilenmeyen), ihbar son günü, yıl dönümü (29 Şubat dahil), periyot dönüşümü,
  maliyet toplamı.
- Uyarılar, okuma hataları ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6098 sayılı Türk Borçlar Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6098.pdf) | Resmî metinle karşılaştırılan maddeler: 100, 115, 182, 315 |

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
