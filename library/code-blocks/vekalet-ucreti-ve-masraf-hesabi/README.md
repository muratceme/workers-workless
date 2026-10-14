# Vekâlet Ücreti ve Masraf Hesabı · Kod Bloğu

> Dava Takip › Hukuk Sekreteri · Workers / Workless

Avukatlık Asgari Ücret Tarifesi'nin (AAÜT) **nispi dilimleriyle** dava dosyaları için karşı tarafa yüklenecek vekâlet ücretini
hesaplar: kabul edilen kısım için davacı, reddedilen kısım için davalı lehine. Yargılama giderlerini **haklılık oranına** göre
paylaştırır. Dilimler tarife dosyasından okunur; yeni tarife yayımlandığında yalnız dosyayı güncellemek yeterlidir. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Nispi ücret | Her dilim kendi oranıyla hesaplanıp toplanır (tarifenin üçüncü kısmı) |
| Maktu taban | Nispi ücret, davanın görüldüğü mahkeme için tarifede yazan maktu ücretin altında kalamaz (`Maktu Ücret` sütunu / `--maktu`) |
| Üst sınır | Hükmedilen ücret, kabul ya da red edilen miktarı geçemez |
| Davacı / davalı lehine | Kabul edilen miktar / reddedilen miktar (dava değeri − kabul) üzerinden ayrı ayrı |
| Yargılama gideri | Davacının yaptığı giderler × kabul oranı davalıya yüklenir, kalanı davacıda kalır (HMK m. 326, haklılık oranında paylaştırma) |

**Örnek tarife (`ornek_veri/tarife_2025_2026.csv`):** 4 Kasım 2025'te Resmî Gazete'de yayımlanan 2025-2026 AAÜT'nin
üçüncü kısmı:

| Dilim | Oran |
|---|---|
| İlk 600.000 TL | %16 |
| Sonraki 600.000 TL | %15 |
| Sonraki 1.200.000 TL | %14 |
| Sonraki 1.200.000 TL | %13 |
| Sonraki 1.800.000 TL | %11 |
| Sonraki 2.400.000 TL | %8 |
| Sonraki 3.000.000 TL | %5 |
| Sonraki 3.600.000 TL | %3 |
| Sonraki 4.200.000 TL | %2 |
| 18.600.000 TL'yi aşan kısım | %1 |

Kontrol: 900.000 TL → 96.000 + 45.000 = **141.000 TL**. 100.000 TL'lik davada nispi ücret 16.000 TL çıkar, ama asliye hukuk
mahkemesi için maktu ücret 45.000 TL olduğundan **45.000 TL** alınır. Her iki örnek de tarifeyi özetleyen hukuk kaynaklarındaki
örneklerle aynıdır.

**Uyarılar:** okunamayan satır (Yüksek) · maktu ücret girilmemiş dosya (taban kontrolü yapılmaz, Orta) · ücretin miktarla
sınırlandığı dosya (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                            # örnek dosyalar, 2025-2026 dilimleri
python main.py --deger 2.500.000 --kabul 1.600.000 --maktu 45.000 --gider 60.000
python main.py --dosyalar dosyalar.xlsx --tarife tarife_2026_2027.csv
```

## Çıktı

`cikti/vekalet_ucreti.xlsx`: `Dosyalar`, `Dilim Detayı` (her taraf için dilim dilim hesap), `Tarife`, `Uyarılar`.

## Dikkat

- **Tarifeyi güncel tutun.** AAÜT her yıl (genellikle ekim–kasım) yeniden yayımlanır ve yıl içinde değişebilir (2025-2026
  tarifesi 2026'da da değiştirildi). Hükmün kurulduğu tarihte yürürlükte olan tarifenin dilimlerini ve mahkemeye göre maktu
  ücretleri Resmî Gazete'den alın.
- **Tarifenin özel hükümleri uygulanmaz.** İcra takipleri, işe iade, tüketici, nafaka, kamulaştırma, manevi tazminat, ıslah,
  feragat / kabul / sulh ile biten davalar, usulden ret, aynı avukatın birden fazla davası gibi durumların kendi kuralları vardır.
  Sonucu avukat kontrol etmelidir.
- **Tutarlara KDV dahil değildir.** Harç tutarlarını yatırılan harç makbuzlarından girin; araç harç hesaplamaz.
- **Örnek veri:** Dosya numaraları ve tutarlar kurgusaldır.

## Testler

- Tarife özetlerindeki örnekler (500.000, 900.000), dilim sınırları, son dilim, maktu taban, miktar sınırı, kısmi kabul ve red.
- Gider paylaşımı, tarife dosyası okuma ve hataları ("kalan" dilim), dosya okuma, tek dosya ve toplu komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
