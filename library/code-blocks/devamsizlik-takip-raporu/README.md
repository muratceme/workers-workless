# Devamsızlık Takip Raporu · Kod Bloğu

> Kayıt Kabul ve Öğrenci İşleri › Öğrenci İşleri Sorumlusu · Workers / Workless

Yoklama kayıtlarından öğrenci bazında **özürlü ve özürsüz devamsızlığı** toplar. Sınıra yaklaşan ve aşan öğrencileri, art arda
özürsüz devamsızlık nedeniyle veliye bildirilmesi gerekenleri listeler. İnternete ve e-Okul'a bağlanmaz; dışa aktarılmış yoklama
listesini okur.

## Nasıl hesaplar?

| Kural | Açıklama |
|---|---|
| Gün | Tam gün 1, yarım gün 0,5; aynı güne yazılan kayıtlar toplanır, en fazla 1 gün sayılır |
| Lise sınırı | Özürsüz **10 gün**, toplam **30 gün** (Ortaöğretim Kurumları Yönetmeliği m. 36). Birinci dereceden yakınını kaybeden öğrenci için toplam 40 gün (özürsüz yine 10): "Toplam Sınır" sütununa 40 yazın |
| İlkokul / ortaokul | Kodda gün sınırı yok; izlemek isterseniz `--ilk-toplam` ve `--ilk-ozursuz` verin |
| Sınıra yaklaştı | Özürsüz ya da toplam devamsızlık sınırın `--esik` (%80) ve üstünde |
| Art arda özürsüz | Art arda iki okul günü tam gün özürsüz (cuma + pazartesi art arda sayılır): veliye bildirim |
| Özürsüz 5. gün | Lisede özürsüz devamsızlığın 5 güne ulaştığı tarih: sınıf tekrarı riski için veli daveti |

Lise kuralları, MEB'in 2023 yönetmelik değişikliği duyurusuna göredir: sınırı aşan öğrenci ders puanlarına bakılmaksızın başarısız
sayılır, veliye yazılı bildirilir; art arda iki gün özürsüz devamsızlık veliye bildirilir ve devamsızlığın 5. gününde veli okula
davet edilir.

**Uyarılar:** sınır aşıldı · okunamayan satır (Yüksek) · sınıra yaklaştı · art arda 2 gün özürsüz (Orta) · özürsüz 5. gün · aynı
güne fazla kayıt (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (lise, 2025-2026)
python main.py --yoklama yoklama.xlsx
python main.py --yoklama yoklama.xlsx --kademe Ortaokul --ilk-toplam 20 --esik 75
```

Kademe sütunu boşsa `--kademe` (varsayılan Lise) kullanılır. Tür boşsa özürsüz, süre boşsa tam gün sayılır.

## Çıktı

`cikti/devamsizlik_raporu.xlsx`: `Öğrenciler` (sınırı aşanlar ve yaklaşanlar üstte), `Sınıflar`, `Uyarılar`.

## Dikkat

- **Resmî kayıt e-Okul'dur.** Rapor yalnız takip ve veli iletişimi içindir; başarısızlık kararı ve resmî bildirimler e-Okul ve okul
  yönetimi üzerinden yürütülür.
- **Mevzuatı doğrulayın.** Lise sınırları yönetmelikten alınmıştır. İlkokul ve ortaokul için kaynaklar farklı bilgiler verdiğinden
  kodda sınır yoktur. Güncel yönetmelik metinlerini, sağlık raporu, tedavi ve özel durum istisnalarını MEB mevzuatından kontrol edin.
- **Çocuklara ait kişisel veri içerir.** Devamsızlık ve sağlık nedenleri KVKK kapsamında özel nitelikli olabilir; raporu yalnız
  yetkili personelle paylaşın.
- **Örnek veri:** Öğrenciler ve yoklamalar kurgusaldır.

## Testler

- Lise özürsüz ve toplam sınırı, 40 günlük özel sınır, ilkokul / ortaokul eşikleri, yaklaşma eşiği.
- Yarım gün, aynı gün fazla kayıt, hafta sonunu atlayan art arda gün, özürsüz 5. gün, sınıf özeti; okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
