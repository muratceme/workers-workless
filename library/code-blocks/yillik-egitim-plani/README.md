# Yıllık Eğitim Planı · Kod Bloğu

> İnsan Kaynakları › Eğitim ve Gelişim Uzmanı · Workers / Workless

Eğitim ihtiyaç listesinden **bütçeye sığan yıllık eğitim planını**, aylık takvimi ve departman özetini oluşturur.
İnternete bağlanmaz.

## Nasıl planlar?

| Adım | Kural |
|---|---|
| Eğitim maliyeti | Katılımcı sayısı × kişi başı maliyet + sabit maliyet |
| Adam-saat | Katılımcı sayısı × süre (saat) |
| Seçim sırası | Önce zorunlu eğitimler; sonra öncelik (1 en yüksek), eşitlikte düşük maliyetli |
| Bütçe | Zorunlular her durumda plana girer; diğerleri kalan bütçe yettiği sürece |
| Ay ataması | "Tercih Ay" doluysa o ay; boşsa adam-saat yükü en az olan ay, yıla yayılarak |
| Kapalı aylar | `--kapali-aylar` ile verilen aylara otomatik atama yapılmaz |

Bütçeye sığmayan bir eğitim atlanır ve sıradaki daha ucuz eğitim denenir; böylece kalan bütçe boşa gitmez.

**Örnek:** Bütçe 100.000. Zorunlular 35.000; öncelik 1 Liderlik 50.000; öncelik 2 İş İngilizcesi 15.000 → toplam
100.000. İleri Excel (20.000) ve Etkili Sunum (18.000) bütçe dışı kalır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, bütçe 100.000, yıl 2027
python main.py --ihtiyaclar egitim_ihtiyaclari.xlsx --butce 250000
python main.py --ihtiyaclar egitim_ihtiyaclari.xlsx --butce 250000 --yil 2027 --kapali-aylar 7,8
```

Zorunlu sütunlar: Eğitim, Katılımcı Sayısı. Diğerleri: Departman, Süre (saat), Kişi Başı Maliyet, Sabit Maliyet,
Öncelik (boşsa 3), Zorunlu (Evet), Tercih Ay (1–12 veya ay adı). `--butce` verilmezse bütün eğitimler plana alınır.

## Çıktı

`cikti/yillik_egitim_plani.xlsx`:

- `Plan`: ay sırasıyla eğitimler; katılımcı, adam-saat, maliyet; tarih ve eğitmen için boş sütun.
- `Aylık Özet`: ay bazında eğitim sayısı, katılımcı, adam-saat ve maliyet.
- `Departman Özeti`: departman bazında toplamlar ve bütçedeki pay.
- `Bütçe Dışı`: plana giremeyen eğitimler ve nedeni.
- `Uyarılar`: bütçeyi aşan zorunlu eğitimler, bütçe dışı kalanlar, kapalı aya denk gelen tercihler.

## Dikkat

- **Zorunluluk sizin girdinizdir.** Kod hangi eğitimin mevzuat gereği zorunlu olduğunu, tekrar süresini ve asgari
  saatini bilmez. İş sağlığı ve güvenliği gibi yasal eğitimleri tehlike sınıfınıza göre siz işaretlersiniz.
- **Ay ataması bir taslaktır.** Eğitmen ve salon uygunluğu, vardiya düzeni, üretim yoğunluğu ve izin dönemleri hesaba
  girmez; "Önerildi" yazan ayları kendi takviminize göre düzeltin.
- **Aynı eğitim birden çok departman için ayrı satırdaysa ayrı oturum sayılır.** Birleştirilecekse tek satır girin;
  sabit maliyet bir kez yazılsın.
- **Maliyet yalnız doğrudan giderdir.** Katılımcıların işten uzak kaldığı sürenin maliyeti, yol ve konaklama dahil
  değildir (kişi başı maliyete ekleyebilirsiniz).
- **Öncelik sırası ihtiyaç analizine dayanmalı.** Kod verilen önceliği uygular, sorgulamaz.
- **Örnek veri:** Eğitimler, departmanlar ve maliyetler kurgusaldır.

## Testler

- Maliyet ve adam-saat hesabı, seçim sırası, bütçe sınırı, atlanıp sonraki ucuz eğitimin alınması.
- Zorunlu eğitimlerin bütçeyi aşması, bütçesiz plan; ay ataması, yıla yayma, kapalı aylar.
- Aylık ve departman özetleri; okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
