# Brüt-Net Maaş Hesaplama · Kod Bloğu

> İnsan Kaynakları › Bordro ve Özlük İşleri Uzmanı · Mali Müşavirlik › Bordro Elemanı · Workers / Workless

Türkiye'nin **2026** parametreleriyle aylık bordro hesaplar: brütten nete, netten brüte, 12 aylık
kümülatif tablo, vergi dilimi geçişleri ve işveren maliyeti. Tek kişi için komut satırından, çok kişi
için Excel/CSV listesinden çalışır. İnternete bağlanmaz.

## Hesap kuralları

| Kalem | Kural (2026) |
|---|---|
| SGK işçi / işsizlik işçi | %14 / %1 — SGK matrahı brüt, üst sınır 297.270,00 TL (asgari ücretin 9 katı) |
| Gelir vergisi | Matrah = brüt − SGK − işsizlik; **kümülatif** matrah üzerinden ücret tarifesi: 190.000 (%15) · 400.000 (%20) · 1.500.000 (%27) · 5.300.000 (%35) · üstü (%40) |
| Asgari ücret GV istisnası | Takvim ayına göre asgari ücret matrahına düşen vergi (Ocak-Haziran 4.211,33 · Temmuz 4.537,75 · Ağustos-Aralık 5.615,10); eksik günde günlük asgari ücret esas alınır |
| Damga vergisi | Brüt × binde 7,59; asgari ücrete isabet eden kısım (250,70 TL) istisna |
| İşveren SGK / işsizlik | %21,75 / %2; teşvik seçeneği: imalat 5 puan, diğer 2 puan |

Parametreler `tr_parametreler.json` dosyasındadır ve her değerin kaynağı yazılıdır. Yeni dönemde
yalnızca bu dosyayı güncellemeniz yeterlidir.

## Doğrulama

Testler şu bağımsız değerlerle birebir karşılaştırır:
asgari ücret neti 28.075,50 TL ve işveren maliyetleri 40.874,63 / 39.223,13 / 40.214,03 TL;
yıllık asgari ücret GV istisnası toplamı 57.881,23 TL; Ocak 2026'da 50.000 TL brütün neti 40.207,53 TL.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py --brut 85000                  # 12 aylık tablo
python main.py --net 60000                   # her ay 60.000 net için gereken brüt
python main.py --brut 85000 --tesvik imalat  # işveren maliyeti teşvikli
python main.py --girdi personel.xlsx          # toplu hesap
python main.py                               # örnek personel listesiyle dener
```

### Toplu hesap girdisi

| Sütun | Zorunlu | Açıklama |
|---|---|---|
| Ad Soyad | Evet | Excel'de sayfa adı olur |
| Brüt **veya** Net | Evet | Aylık tutar; Net verilirse netten brüte hesaplanır |
| Gün | Hayır | SGK gün sayısı (varsayılan 30) |
| Başlangıç Ayı | Hayır | Yıl içinde işe başlayanlar için (1-12) |
| Önceki Kümülatif Matrah | Hayır | Aynı işverende önceki ayların GV matrahı toplamı |

## Kapsam dışı

Engellilik indirimi, BES kesintisi, ayni yardımlar, yemek/yol istisnaları, sendika aidatı, icra
kesintisi, SGK destek primi (emekli çalışan) ve özel teşvikler hesaba katılmaz. Sonuçlar bilgilendirme
amaçlıdır; resmî bordro için bordro programınızı esas alın.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
