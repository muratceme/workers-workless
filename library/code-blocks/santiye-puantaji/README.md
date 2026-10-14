# Şantiye Puantajı · Kod Bloğu

> Şantiye › Saha Mühendisi · Workers / Workless

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

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
