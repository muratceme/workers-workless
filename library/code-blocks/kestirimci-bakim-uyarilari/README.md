# Kestirimci Bakım Uyarıları · Kod Bloğu

> Üretim › Bakım Onarım › Bakım Mühendisi · Workers / Workless

Titreşim, sıcaklık ve akım ölçümlerini ekipman limitleriyle karşılaştırır. Arıza riski yüksek ekipmanları risk ve
kritiklik sırasıyla listeler. Ölçümler portatif cihaz rotasından ya da sensör / SCADA dışa aktarımından gelebilir.
İnternete bağlanmaz.

## Ne kontrol eder?

| Kontrol | Önem |
|---|---|
| Son ölçüm alarm seviyesinde | Kritik |
| Son ölçüm uyarı seviyesinde | Yüksek |
| Eğilim: son 30 günün doğrusal eğilimiyle alarm seviyesine 30 gün içinde ulaşılıyor (`--pencere`, `--ufuk`) | Yüksek |
| Eğilim: uyarı seviyesine 30 gün içinde ulaşılıyor | Orta |
| Aynı ekipmanda iki ayrı parametrede uyarı / alarm veya hızlı eğilim (çoklu belirti) | Yüksek |
| Ani değişim (titreşim): baz çizgisinden sapma, B/C sınırının %25'inden büyük | Orta |
| Sabit değer: art arda 6 kez aynı değer (`--tekrar`); yalnız serinin geri kalanı ölçümden ölçüme belirgin değişiyorsa | Orta |
| Ölçüm gecikmiş (periyodun 1,5 katı), hiç ölçüm yok, ekipman tablosunda olmayan ekipman | Orta |
| Geçmiş aşım (son ölçüm normale dönmüş), sıfır / negatif okuma, limit tanımsız | Bilgi |

**Eşik ve eğilim kuralları:**
- **Eğilim:** Pencerede en az 5 geçerli ölçüm ve R² ≥ 0,6 gerekir. Sıfır ve negatif okumalar (ör. makine duruşta)
  eğilim hesabına alınmaz.
- **Baz çizgisi:** Serinin ilk 5 geçerli ölçümünün medyanıdır.
- **Ani değişim:** ISO 10816 / 20816 serisinin "Kriter II" yaklaşımına dayanır: seviye kabul bölgesinde olsa da
  titreşimdeki belirgin değişim araştırılmalıdır.

**Limitler:**
- **Titreşim:** Ekipman tablosunda limit yoksa ISO 10816-3 sınıfına göre B/C sınırı uyarı, C/D sınırı alarm olarak
  kullanılır. Değerler mm/s RMS'dir.

| Sınıf | A/B | B/C (uyarı) | C/D (alarm) |
|---|---|---|---|
| Grup 1 rijit (300 kW – 50 MW) | 2,3 | 4,5 | 7,1 |
| Grup 1 esnek | 3,5 | 7,1 | 11,0 |
| Grup 2 rijit (15 – 300 kW) | 1,4 | 2,8 | 4,5 |
| Grup 2 esnek | 2,3 | 4,5 | 7,1 |

- **Akım:** Limit yoksa plaka (nominal) akımı alarm sayılır. Uyarı, nominal × 0,90'dır (`--akim-uyari-orani`).
- **Sıcaklık:** Yalnız ekipman tablosundaki limitler kullanılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek tesis, 8 ekipman, 90 günlük ölçüm
python main.py --ekipmanlar ekipmanlar.xlsx --olcumler olcumler.xlsx
python main.py --ekipmanlar e.xlsx --olcumler o.xlsx --rapor-tarihi 30.09.2026 --pencere 45 --ufuk 21
```

| Dosya | Sütunlar |
|---|---|
| Ekipman listesi | Ekipman Kodu, Ekipman Adı, Kritiklik, ISO 10816-3 Sınıfı (ör. "Grup 2 rijit"), Titreşim Uyarı, Titreşim Alarm (mm/s), Sıcaklık Uyarı, Sıcaklık Alarm (°C), Nominal Akım, Akım Uyarı, Akım Alarm (A), Ölçüm Periyodu (gün) |
| Ölçümler | Tarih, Saat (isteğe bağlı), Ekipman Kodu, Ölçüm Noktası, Parametre (Titreşim / Sıcaklık / Akım), Değer |

Parametre adı "Titreşim hızı RMS", "Yatak sıcaklığı", "Akım" gibi yazılabilir. Aynı ekipmanın farklı ölçüm noktaları
(Motor DE, NDE, fan yatağı) ayrı seriler olarak değerlendirilir.

## Çıktı

`kestirimci_bakim_uyarilari.xlsx`:
- `Ekipman Riski`: risk ve kritiklik sırası, bulgu özeti, parametreye göre önerilen inceleme, boş
  "Planlanan Müdahale / Karar" sütunu.
- `Bulgular`: önem, tür, açıklama ve "İnceleme Sonucu" sütunu.
- `Seriler`: nokta ve parametre bazında uyarı / alarm limiti, baz çizgisi, son değer, haftalık eğim, R² ve uyarıya /
  alarma kalan tahmini gün.
- `Grafikler`: bulgusu olan serilerin limit çizgili grafikleri.
- `Ölçümler` ve `Ekipmanlar`. `Ekipmanlar` sayfası uygulanan limitleri ve kaynaklarını gösterir.

## Dikkat

- **Limitleri kontrol edin:** ISO 10816-3 değerleri genel bir başlangıç noktasıdır. Ölçümün standarttaki koşullarla
  yapılması gerekir: yatak gövdesinden, 10–1000 Hz aralığında. Ayrı tahrikli pompalar (ISO 10816-7), pistonlu
  makineler, dişli kutuları ve fanlar için ilgili standart bölümünün veya üreticinin değerlerini girin. Standardın
  güncel sürümüyle (ISO 20816-3) karşılaştırın.
- **Örnek değerler:** Sıcaklık limitleri ve akım uyarı oranı örnektir. Yatak tipi, yağlama, motor yalıtım sınıfı ve
  üretici önerisine göre girin.
- **Eğilim tahmini:** Tahmin doğrusal varsayıma dayanır. Rulman hasarı gibi arızalar son aşamada hızlanabilir, bu
  yüzden "kalan gün" garanti değildir. Kaynak teşhisi için spektrum / zarf analizi, termografi ve yağ analizi gibi
  yöntemler gerekir. Durdurma ve müdahale kararı bakım mühendisinindir.
- **Örnek veri:** Örnek tesis, ekipmanlar ve ölçümler kurgusaldır.

## Testler

Şunlar test edilir:
- ISO ve nominal akımdan türetilen limitler.
- Alarm, uyarı ve eğilim bulguları (kalan gün); çoklu belirti; ani değişim.
- Çözünürlüğe duyarlı sabit değer kuralı; gecikmiş ve eksik ölçüm; geçmiş aşım; sıfır okuma.
- Risk sıralaması, Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
