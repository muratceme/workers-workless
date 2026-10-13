# Acente Kontenjan (Allotment) Takibi · Kod Bloğu

> Rezervasyon ve Gelir Yönetimi › Rezervasyon Sorumlusu · Workers / Workless

Tur operatörü ve acente kontratlarındaki **kontenjanı, release tarihini ve kullanımı** konaklama günü bazında izler;
aşımları ve release'i yaklaşan boş kontenjanı listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| Kullanılan | O gece konaklayan (giriş ≤ gün < çıkış), iptal olmayan rezervasyonların oda sayısı |
| Kalan | Kontenjan − kullanılan |
| Release tarihi | Konaklama günü − release gün sayısı |
| Release oldu | Bugün ≥ release tarihi: kullanılmayan kontenjan otele dönmüş sayılır |
| Release yaklaşıyor | Release tarihine `--yaklasan` (7) gün veya daha az var ve boş kontenjan var |
| Aşım | Kullanılan > kontenjan |

Özet, bugünden itibaren `--ufuk` (30) gün içindeki **oda-gece** toplamlarıdır: kontenjan, kullanılan, kullanım %,
release ile dönen, release'i yaklaşan boş, hâlâ satılabilir kalan ve aşım.

**Örnek:** Kontenjan 10 oda, release 7 gün; bugün 12.10. 19.10 gecesinin release tarihi 12.10 → "Release oldu".
20.10 gecesinin release tarihi 13.10 → 1 gün kaldı, "Release yaklaşıyor". 16.10 gecesi 6 + 5 = 11 oda → 1 oda aşım.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, bugün 12.10.2026
python main.py --kontratlar kontratlar.xlsx --rezervasyonlar rezervasyonlar.xlsx
python main.py --kontratlar kontratlar.xlsx --rezervasyonlar rezervasyonlar.xlsx --tarih 12.10.2026 --ufuk 60 --yaklasan 10
```

- Kontratlar: Acente, Oda Tipi, Başlangıç, Bitiş, Kontenjan (günlük oda), Release (gün; boşsa 0).
- Rezervasyonlar: Acente, Oda Tipi, Giriş, Çıkış, Oda Sayısı (boşsa 1), Durum ("İptal" ile başlayanlar sayılmaz).

Acente ve oda tipi adları iki dosyada aynı yazılmalıdır (büyük-küçük harf ve Türkçe karakter farkı sorun olmaz).

## Çıktı

`cikti/kontenjan_takibi.xlsx`:

- `Özet`: acente ve oda tipi bazında oda-gece toplamları.
- `Günlük Durum`: her konaklama günü için kontenjan, kullanılan, kalan, release tarihi, durum.
- `Yaklaşan Release`: release tarihine göre sıralı, boş kontenjanı olan günler.
- `Uyarılar`: kontenjan aşımı, yaklaşan release, kontenjan dışı rezervasyon, çakışan kontrat satırları.

## Dikkat

- **Release tanımı kontrata göre değişir.** Kod "konaklama gününden N gün önce, o günün başında" kabul eder. Kontratınız
  release gününü dahil sayıyorsa veya saat belirtiyorsa bir gün fark çıkabilir; kontratı esas alın.
- **Kontenjan otomatik kapanmaz.** Kod rapor üretir; satışı kapatma (stop sale) ve odaları geri alma işlemini PMS veya
  kanal yöneticinizde siz yaparsınız.
- **Garantili kontenjan ve taahhüt ayrı izlenir.** Release'siz garantili odalar (commitment) için release sütununu 0
  girin; boş kalan odanın faturalanması kontrat konusudur.
- **Asgari konaklama, stop sale ve erken rezervasyon koşulları kapsam dışıdır.**
- **Aynı günleri kapsayan birden çok kontrat satırı toplanır.** Dönem geçişlerinde tarihlerin çakışmadığını kontrol edin.
- **Örnek veri:** Acenteler, kontratlar ve rezervasyonlar kurgusaldır.

## Testler

- Günlük kullanılan, kalan, release tarihi ve durumlar; özet oda-gece toplamları.
- `--ufuk` ve `--yaklasan` ayarları, release 0, ufuk dışında kalan kontrat.
- Birden çok acente ve oda tipi, çakışan kontrat, kontenjan dışı rezervasyon, iptal; okuma ve komut satırı hataları.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
