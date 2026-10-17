# Şantiye Puantajı · Kod Bloğu

> Şantiye › Saha Mühendisi · Gerisi Bende

Kendi personelinin ve taşeronların günlük çalışma kayıtlarından **kişi × gün puantaj cetveli** çıkarır; **firma, ekip ve
imalat** bazında saat ve adam-gün toplar; günlük ve haftalık süre aşımlarını ve aynı kişinin aynı gün iki firmada
yazılmasını işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Çalışma saati | Saat sütunu; yoksa Gün sütunu (1, 0,5) × `--gun-saat` (varsayılan 7,5); fazla mesai sütunu eklenir |
| Adam-gün | Çalışma saati ÷ `--gun-saat` |

Kişi, sicil numarası yazılmışsa sicilden, yoksa addan tanınır.

| Kontrol | Kural |
|---|---|
| Günlük sınır | Kişinin bir gündeki toplam çalışması `--gunluk-sinir` (varsayılan 11) saati aşıyor |
| Haftalık süre | Kişinin bir haftadaki (pazartesi-pazar) toplamı `--haftalik` (varsayılan 45) saati aşıyor (bilgi) |
| İki firma | Aynı kişi aynı gün iki firmanın puantajında |

4857 sayılı İş Kanunu'na göre haftalık çalışma süresi en çok 45 saattir ve günlük çalışma 11 saati aşamaz; varsayılan
eşikler buna göre seçilmiştir. Fazla çalışma ücreti ve yıllık sınırı hesaplanmaz.

**Örnek:** Demir ekibinden Kemal A. 1 Ekim perşembe 7,5 saat + 4 saat fazla mesai = **11,5 saat** → günlük sınır uyarısı.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                         # örnek veri (iki hafta)
python main.py --puantaj gunluk_puantaj.xlsx
python main.py --puantaj gunluk_puantaj.xlsx --gun-saat 9 --haftalik 45
```

Her satır bir kişinin bir gündeki bir işidir (aynı gün iki imalatta çalışan kişi iki satır olabilir). Zorunlu sütunlar:
Tarih, Personel ve Saat ya da Gün.

## Çıktı

`cikti/santiye_puantaji.xlsx`:

- `Puantaj Cetveli`: kişi × gün saatleri, çalışılan gün, toplam saat, fazla mesai, adam-gün; sınırı aşan günler
  kırmızı, pazar çalışmaları sarı.
- `Özet`: firma, ekip ve imalat bazında kişi, gün, saat, adam-gün.
- `Günlük`: gün bazında şantiyedeki kişi sayısı ve saat.
- `Uyarılar`.

## Dikkat

- **Bordro ve SGK bildirimi değildir.** Prim günü, ücret, fazla çalışma ücreti ve hafta tatili hesabı bordro biriminde
  yapılır; taşeron personelinin SGK bildirimleri taşeronun sorumluluğundadır, asıl işveren olarak ayrıca kontrol edin.
- **Kişisel veri içerir.** Ad, sicil ve çalışma saatleri çalışan verisidir; erişimi sınırlayın.
- **Adam-gün katsayısı politikadır.** 7,5 saat (45 saat ÷ 6 gün) varsayılandır; şantiyenizin çalışma düzenine göre
  değiştirin.
- **Örnek veri:** Kişi adları, firmalar ve saatler kurgusaldır.

## Testler

- Saat ve gün girişinden çalışma saati, fazla mesai, adam-gün, kişi, firma, ekip, imalat ve gün toplamları.
- Günlük ve haftalık sınır, aynı gün iki firma, seçenekler, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |

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
