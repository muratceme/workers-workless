# SKS Öz Değerlendirme Takibi · Kod Bloğu

> Kalite Yönetimi › Kalite Birim Sorumlusu · Workers / Workless

Sağlıkta Kalite Standartları (SKS) **öz değerlendirme sonuçlarını** birim ve bölüm bazında puanlar; kısmen ya da hiç
karşılanmayan standartları ve **eksik / güncel olmayan dokümanları** sorumlu ve terminle aksiyon listesine çevirir.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Kazanılan puan | Karşılıyor: standart puanı · Kısmen: puan × `--kismen` (varsayılan %50) · Karşılamıyor: 0 · Değerlendirme dışı: hesaba girmez |
| Uyum oranı % | Σ kazanılan ÷ Σ değerlendirilen standart puanı × 100 (birim, bölüm, toplam) |
| Eksik doküman | Doküman Durumu "Yok" ya da "Güncel değil" |
| Aksiyon | Kısmen / Karşılamıyor değerlendirilen standart ya da eksik doküman |

Puan sütunu boşsa her standart 1 puan sayılır. Değerlendirme "Evet / Hayır / Kısmen / DD" yazımlarıyla da girilebilir.

**Uyarılar:** birim uyum oranı `--hedef` (varsayılan %80) altında · termini geçmiş aksiyon · sorumlusu yazılmamış aksiyon.

**Örnek:** Acil serviste değerlendirilen standartların toplam puanı 125, kazanılan 95 → uyum **%76** (hedef %80 →
uyarı).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                         # örnek veri (30.09.2026)
python main.py --degerlendirme oz_degerlendirme.xlsx --tarih 30.09.2026
python main.py --degerlendirme oz_degerlendirme.xlsx --kismen 0 --hedef 90
```

Zorunlu sütunlar: Standart Kodu ve Değerlendirme. Her satır bir birimin bir standarttaki öz değerlendirmesidir.

## Çıktı

`cikti/sks_oz_degerlendirme.xlsx`:

- `Birimler`: uyum oranına göre (en düşük önce) sıralı; karşılıyor / kısmen / karşılamıyor sayıları, puanlar, eksik
  doküman.
- `Bölümler`: aynı göstergeler bölüm bazında.
- `Aksiyon Listesi`: termine göre sıralı; gecikme gün sayısı, neden, gerekli doküman, sorumlu.
- `Tüm Değerlendirmeler`, `Uyarılar`.

## Dikkat

- **Resmî değerlendirme puanı değildir.** Puanlama şeması (standart puanları, kısmen karşılamanın değeri) kurumunuzun
  öz değerlendirme yöntemine göre girilmelidir; Sağlık Bakanlığı değerlendirmesinin yöntemini ve puanını taklit etmez.
- **Standartları güncel setten alın.** SKS setleri ve standart kodları sürüm sürüm değişir; standart listesini Bakanlığın
  yayımladığı güncel setten hazırlayın. Örnek verideki kodlar ("ÖD-1.01") kurgusaldır.
- **Kanıt değerlendirmesi insan işidir.** Dokümanın var olması uygulamanın yapıldığını göstermez; kayıt ve saha
  gözlemiyle doğrulayın.
- **Örnek veri:** Birimler, standart kodları, puanlar ve terminler kurgusaldır.

## Testler

- Karşılıyor / kısmen / karşılamıyor / değerlendirme dışı puanlaması, birim, bölüm ve toplam uyum oranı.
- Eksik doküman, aksiyon listesi ve sırası, geciken termin, sorumlu yok uyarısı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
