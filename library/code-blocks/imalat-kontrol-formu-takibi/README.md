# İmalat Kontrol Formu Takibi · Kod Bloğu

> Kalite Kontrol › Kalite Kontrol Mühendisi · Workers / Workless

İmalat kontrol taleplerini (muayene ve test planı / ITP adımları) **blok · kat · imalat** bazında izler; her konumun
güncel onay durumunu bir **durum matrisinde** gösterir, red oranını ve karar süresini hesaplar, bekleyen ve
kapatılmamış kontrolleri listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| Güncel durum | Her blok · kat · imalat (· kontrol noktası) için en son talebin durumu; ikinci ve sonraki talepler "n. başvuru" olarak gösterilir |
| Red oranı % | Reddedilen ÷ karar verilen (onay + şartlı onay + red) × 100 |
| Karar süresi | Karar tarihi − talep tarihi (gün) |

Durum "Onaylandı / Uygun", "Şartlı onay", "Reddedildi / Uygun değil", "Bekliyor / Açık" yazımlarıyla tanınır; boş
durum bekliyor sayılır.

**Uyarılar:** `--gun` (varsayılan 2) günden uzun bekleyen talep · son durumu red olup yeniden başvurulmamış imalat ·
şartlı onayı kapatılmamış imalat (bilgi).

**Örnek:** 40 talepten 38'ine karar verilmiş, 8'i reddedilmiş → red oranı **%21,05**; A Blok 3. kattaki "beton döküm
öncesi" kontrolü reddedilmiş ve yeniden başvuru yok.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                           # örnek veri (10.10.2026)
python main.py --formlar kontrol_formlari.xlsx --tarih 10.10.2026
python main.py --formlar kontrol_formlari.xlsx --gun 1
```

Zorunlu sütunlar: Talep Tarihi ve İmalat Türü. Aynı imalatta birden çok kontrol noktası izleniyorsa Kontrol Noktası
sütununu doldurun. Kat değerleri doğal sırayla dizilir (1, 2, 10).

## Çıktı

`cikti/imalat_kontrol_takibi.xlsx`:

- `Durum Matrisi`: blok · kat satırları, imalat sütunları; hücrede güncel durum (renkli).
- `Özet`: blok, kat, imalat ve toplam bazında talep, onay, şartlı onay, red, bekleyen, red oranı, karar süresi.
- `Bekleyenler`, `Tüm Formlar`, `Uyarılar`.

## Dikkat

- **Kontrol planı projeye özeldir.** Hangi imalatın hangi aşamada kontrol edileceği, kimin onaylayacağı ve bekleme süresi
  sözleşme ve kalite planına göre belirlenir; kod yalnız kayıtları izler.
- **Onay, imzalı kontrol formunun yerine geçmez.** Tablo, form arşivinin takip listesidir.
- **Red oranı tek başına kalite göstergesi değildir;** kontrolün titizliği ve kayıt alışkanlığı oranı etkiler.
- **Örnek veri:** Bloklar, katlar, form numaraları ve kişiler kurgusaldır.

## Testler

- En son talep ve başvuru sayısı, durum matrisi, red oranı, karar süresi, kırılımlar.
- Bekleyen, reddedilmiş ve şartlı onaylı kontrol uyarıları, durum tanıma, doğal kat sırası, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
