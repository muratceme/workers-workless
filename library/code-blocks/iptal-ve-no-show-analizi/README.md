# İptal ve No-Show Analizi · Kod Bloğu

> Ön Büro › Ön Büro Müdürü · Workers / Workless

Otel rezervasyonlarının **iptal ve gelmeme (no-show) oranlarını** kanal, oda tipi, rezervasyon öncesi süre ve giriş ayı
bazında hesaplar; kayıp oda-gece ve geliri, iptallerin girişe ne kadar kala yapıldığını gösterir. İnternete ve PMS'e
bağlanmaz; dışa aktarılmış rezervasyon listesini okur.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| İptal oranı % | İptal edilen ÷ toplam rezervasyon × 100 |
| No-show oranı % | Gelmeyen ÷ iptal edilmemiş (gelmesi beklenen) rezervasyon × 100 |
| Oda-gece | Çıkış − giriş (ya da Gece sütunu) |
| Kayıp oda-gece / gelir | İptal ve no-show rezervasyonların oda-gecesi ve tutarı |
| Rezervasyon öncesi süre | Giriş − rezervasyon tarihi: 0–7 · 8–30 · 31–90 · 90+ gün |
| İptal zamanı | Giriş − iptal tarihi: aynı gün · 1–3 · 4–7 · 8–30 · 30+ gün |

Durum metinden tanınır: Konakladı / Gerçekleşti / Check-in / Check-out · İptal / Cancelled · No-show / Gelmedi.

**Uyarılar:** en az `--min` (varsayılan 10) rezervasyonlu kanal ya da oda tipinde iptal oranı `--iptal` (varsayılan %25)
ya da no-show oranı `--noshow` (varsayılan %5) üstünde · son dakika iptallerinin payı (bilgi) · tarihi yazılmamış iptaller.

**Örnek:** OTA A kanalından 84 rezervasyonun 29'u iptal → **%34,52**; iptal edilmeyen 55 rezervasyondan 5'i gelmedi →
no-show **%9,09**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (eylül 2026)
python main.py --rezervasyonlar rezervasyonlar.xlsx
python main.py --rezervasyonlar rezervasyonlar.xlsx --min 20 --iptal 20 --noshow 3
```

Zorunlu sütunlar: Giriş Tarihi ve Durum. Rezervasyon tarihi, çıkış, iptal tarihi ve tutar isteğe bağlıdır; yazılmayan
göstergeler hesaplanmaz.

## Çıktı

`cikti/iptal_noshow_analizi.xlsx`:

- `Kırılımlar`: toplam, kanal, oda tipi, rezervasyon öncesi süre ve giriş ayı bazında rezervasyon, iptal, no-show, oranlar,
  oda-gece, kayıp oda-gece, kayıp gelir, ortalama rezervasyon öncesi süre.
- `İptal Zamanı`: iptallerin girişe kalan süreye göre dağılımı.
- `Uyarılar`.

## Dikkat

- **Kayıp gelir brüttür.** Tahsil edilen iptal ve no-show ücretleri, ön ödemeler ve yeniden satılan odalar düşülmez;
  gerçek kayıp daha düşük olabilir.
- **No-show tanımı otelden otele değişir.** Geç iptal ya da kısa konaklama (erken çıkış) bu analizde yoktur; PMS'teki
  durum kodlarınızı tanınan adlara çevirin.
- **Overbooking ve iptal politikası** kararları için tek dönem yetersizdir; sezon ve özel günleri ayrı değerlendirin.
- **Eşikler politika varsayılanıdır;** pazarınıza ve kanal yapınıza göre değiştirin.
- **Örnek veri:** Kanallar, rezervasyonlar ve tutarlar kurgusaldır.

## Testler

- İptal ve no-show oranı, oda-gece, kayıp oda-gece ve gelir, ortalama rezervasyon öncesi süre.
- Kanal / oda tipi / süre bantları, iptal zamanı dağılımı, eşik uyarıları, durum tanıma, okuma ve komut satırı.

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
