# Test Sonuçları Spesifikasyon Karşılaştırması · Kod Bloğu

> Ar-Ge ve Ürün Geliştirme › Ar-Ge Mühendisi · Workers / Workless

Prototip test sonuçlarını **spesifikasyon limitleriyle** karşılaştırır; uygunsuz sonuçları ve **prototipler arası
eğilimi** (limite yaklaşıyor mu, uzaklaşıyor mu) raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Uygun   = alt limit ≤ sonuç ≤ üst limit         (limitin kendisi uygundur; yalnız alt veya yalnız üst limit olabilir)
Marj    = ortalamanın en yakın limite uzaklığı   (ortalama limit dışındaysa eksi)
Marj %  = marj ÷ (üst − alt) × 100               tek taraflı limitte: marj ÷ |limit| × 100
Eğilim  = son prototipin marj %'si − ilk prototipin marj %'si
```

| Eğilim | Koşul |
|---|---|
| İyileşiyor | Fark > `--egilim-esik` (varsayılan 1 puan) |
| Kötüleşiyor | Fark < −`--egilim-esik` |
| Değişmedi | Arada |
| Tek prototip | Parametrenin yalnız bir prototipte sonucu var |

**Dikkat** işareti iki durumda konur: son prototipte uygunsuz sonuç varsa; veya eğilim kötüleşiyor ve son marj
`--marj-esik` (varsayılan %10) altındaysa.

**Örnek:** sıcaklık artışı üst limiti 35 °C. Ortalamalar P1 30,5 → P2 32,5 → P3 35,0; marj %12,86 → %7,14 → %0,00.
Sonuçlar hâlâ limit civarında ama eğilim kötüleşiyor ve marj kalmamış: "Dikkat".

Prototip sırası dosyada ilk görüldükleri sıradır. Sayılarda virgül de nokta da ondalık ayırıcıdır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --sonuclar test_sonuclari.xlsx --spesifikasyon spesifikasyon.xlsx --marj-esik 15
```

| Dosya | Sütunlar |
|---|---|
| Test sonuçları | Parametre, Sonuç (zorunlu); Prototip, Numune, Tarih, Koşul |
| Spesifikasyon | Parametre (zorunlu); Alt Limit, Üst Limit (en az biri), Hedef, Birim |

## Çıktı

`cikti/test_spesifikasyon.xlsx`:

- `Prototip Özeti`: prototip bazında uygunsuz sonuç sayısı ve uygunsuz parametreler.
- `Eğilim`: parametre bazında prototip ortalamaları, ilk ve son marj, eğilim ve dikkat notu; boş "Yorum" sütunu.
- `Parametre Özeti`: parametre × prototip; ortalama, en küçük, en büyük, marj, hedeften fark.
- `Sonuçlar`: her test sonucu ve değerlendirmesi.
- `Uyarılar`: spesifikasyonu olmayan veya hiç test edilmemiş parametreler, okunamayan değerler.

## Dikkat

- **Eğilim istatistiksel test değildir.** Yalnız ilk ve son prototipin ortalamaları karşılaştırılır; birkaç numuneyle
  görülen fark ölçüm değişkenliğinden kaynaklanıyor olabilir.
- **Test koşulları aynı olmalı.** Farklı koşullarda (sıcaklık, gerilim, yük) alınan sonuçları aynı parametre adıyla
  karıştırmayın; koşulu parametre adına ekleyin (ör. "Debi @ 12 V").
- **Ölçüm belirsizliği hesaba katılmaz.** Limite çok yakın sonuçlarda karar kuralınızı ayrıca uygulayın.
- **Onay kararı değildir.** Tasarım doğrulama ve tip onayı kararları ilgili standardın numune sayısı ve kabul
  koşullarına göre verilir.
- **Örnek veri:** Prototipler, parametreler ve sonuçlar kurgusaldır.

## Testler

- Uygunluk sınırları; çift ve tek taraflı limitte marj (limit dışı ortalama dahil).
- Eğilim yönleri, eşikler, dikkat koşulları; parametre ve prototip özetleri.
- Spesifikasyonu olmayan ve test edilmemiş parametreler; okuma hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
