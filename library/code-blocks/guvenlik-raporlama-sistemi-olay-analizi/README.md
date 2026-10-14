# Güvenlik Raporlama Sistemi Olay Analizi · Kod Bloğu

> Kalite Yönetimi › Kalite Birim Sorumlusu · Workers / Workless

Hasta ve çalışan güvenliği **olay bildirimlerini** tür, kategori, birim, zarar düzeyi, ay ve kök neden bazında
analiz eder; **ramak kala** ve **önlenebilir** olay oranlarını, ciddi sonuçlu olayları ve uzun süredir kapanmayan
bildirimleri gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Dağılım ve pay | Kırılım başına bildirim ÷ toplam bildirim × 100 |
| Ramak kala oranı % | Ramak kala bildirimi ÷ toplam bildirim × 100 |
| Önlenebilir oranı % | Önlenebilir ÷ önlenebilirliği değerlendirilmiş bildirim × 100 |
| 1000 hasta gününe bildirim | Birim bildirimi ÷ birim hasta günü × 1000 (`--hasta-gunu` verilirse) |

Zarar düzeyi metinden tanınır: Ramak kala · Zarar yok · Hafif · Orta · Ciddi (ağır, kalıcı) · Ölüm. Olay türü "hasta"
ya da "çalışan / personel" sözcüğünden ayrılır; diğer yazımlar kendi adıyla kalır.

**Uyarılar:** ciddi zarar ya da ölümle sonuçlanan her olay (Yüksek) · bu olaylarda DÖF numarası yoksa (Yüksek) ·
`--gun` (varsayılan 30) günden uzun süredir kapanmamış bildirim (Orta) · önlenebilirliği değerlendirilmemiş bildirim
sayısı (Bilgi).

**Örnek:** 74 bildirimin 21'i ramak kala → **%28,38**. Önlenebilirliği değerlendirilen 58 bildirimin 42'si önlenebilir
→ **%72,41**. Yoğun bakımda 1.100 hasta gününde 18 bildirim → **16,36** / 1000 hasta günü.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                 # örnek veri (30.09.2026)
python main.py --bildirimler olay_bildirimleri.xlsx --tarih 30.09.2026
python main.py --bildirimler olay_bildirimleri.xlsx --hasta-gunu birim_hasta_gunu.xlsx --gun 15
```

Zorunlu sütunlar: Olay Tarihi ve Birim. Diğer sütunlar boşsa "Belirtilmemiş" sayılır.

## Çıktı

`cikti/grs_olay_analizi.xlsx`:

- `Özet`: bildirim, ramak kala, ciddi, önlenebilir ve açık bildirim sayı ve oranları.
- `Dağılımlar`: olay türü, zarar düzeyi, kategori, birim (hasta gününe oran), ay ve kök neden kırılımları.
- `Açık Bildirimler`: en eskiden yeniye.
- `Bildirimler`, `Uyarılar`.

## Dikkat

- **Bildirim sayısı güvenlik düzeyini göstermez.** Çok bildirim yapan birim daha güvensiz değil, bildirim kültürü daha
  gelişmiş olabilir; birimleri sayı üzerinden cezalandırıcı biçimde karşılaştırmayın. GRS bildirimleri suçlama amacıyla
  kullanılmamalıdır.
- **Önlenebilirlik ve zarar düzeyi** değerlendirme ekibinin kararıdır; kod yalnız yazılanı sayar.
- **Yasal ve kurumsal yükümlülükler:** Ciddi sonuçlu olaylarda kök neden analizi, bakanlık bildirimleri ve hasta
  hakları süreçleri kurum prosedürüne ve güncel mevzuata göre yürütülür; rapor bunların yerine geçmez.
- **Kişisel ve sağlık verisi** içerebilir; analiz dosyasına hasta ve çalışan kimliği koymayın.
- **Örnek veri:** Bildirimler, birimler ve hasta günleri kurgusaldır.

## Testler

- Ramak kala ve önlenebilir oranı, birim / zarar / kategori / ay dağılımı, hasta gününe oran.
- Ciddi olay ve DÖF uyarısı, açık bildirim yaşı, metin tanıma, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
