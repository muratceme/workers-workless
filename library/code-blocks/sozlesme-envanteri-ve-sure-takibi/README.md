# Sözleşme Envanteri ve Süre Takibi · Kod Bloğu

> Hukuk › Hukuk Müdürü · Workers / Workless

Sözleşme listenizi süre takvimine çevirir. Asıl sorduğu soru şudur: **"Bu sözleşme için en geç hangi gün karar
vermem gerekiyor?"** Otomatik yenilenen sözleşmelerde bu gün bitiş değil, fesih bildirim süresinin son günüdür;
kaçırılırsa sözleşme bir dönem daha yenilenir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Güncel bitiş | Otomatik yenilenen sözleşmenin yazılı bitişi geçmişse, yenileme süresi (boşsa 12 ay) eklenerek durum tarihinden sonrasına taşınır |
| Fesih bildirim son günü | Güncel bitiş − fesih bildirim süresi (takvim günü) |
| Karar tarihi | Otomatik yenilenenlerde fesih bildirim son günü; diğerlerinde bitiş |

| Durum | Anlamı |
|---|---|
| Süresi dolmuş | Otomatik yenilenmeyen sözleşmenin bitişi geçmiş |
| Bildirim süresi geçti | Otomatik yenilenen sözleşmede bildirim son günü geçmiş; sözleşme yenilenecek |
| Fesih bildirimi yaklaşıyor | Bildirim son gününe `--gun` (90) gün veya daha az kalmış |
| Bitiş yaklaşıyor | Otomatik yenilenmeyen sözleşmenin bitişine `--gun` gün veya daha az kalmış |
| Yürürlükte, Süresiz | Karar tarihi uzak; bitiş tarihi olmayan sözleşme |

**Örnek:** Bitiş 30.11.2026, otomatik yenileme var, fesih bildirim süresi 30 gün → bildirim için son gün
**31.10.2026**. Bu tarihe kadar bildirim yapılmazsa sözleşme 12 ay daha yenilenir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri, durum tarihi 12.10.2026
python main.py --sozlesmeler sozlesmeler.xlsx
python main.py --sozlesmeler sozlesmeler.xlsx --tarih 12.10.2026 --gun 120
```

Zorunlu sütunlar: Karşı Taraf, Bitiş (süresiz sözleşmede boş bırakılır). Diğerleri: Sözleşme No, Konu, Tür,
Başlangıç, Otomatik Yenileme (Evet / Hayır), Yenileme Süresi (ay), Fesih Bildirim Süresi (gün), Sorumlu, Yıllık Tutar.

## Çıktı

`cikti/sozlesme_envanteri.xlsx`:

- `Envanter`: durum, yazılı ve güncel bitiş, fesih bildirim son günü, karar tarihi, kalan gün. Karar bekleyenler
  en üstte, tarihi en yakın olan önce.
- `Tür Özeti`, `Sorumlu Özeti`: sözleşme sayısı, karar bekleyen sayısı, yıllık tutar.
- `Uyarılar`: süresi dolanlar, yaklaşan bildirim ve bitiş tarihleri, yenilenmiş sayılanlar, eksik bilgiler.

## Dikkat

- **Sözleşme metni esastır.** Kod yalnız tabloya yazdığınız tarih ve süreleri kullanır. Bildirimin şekli (noter,
  iadeli taahhütlü, kayıtlı elektronik posta), sürenin gönderim mi tebliğ mi üzerinden sayılacağı ve son günün
  tatile denk gelmesi hâli sözleşmeye ve mevzuata göre değişir; son günü beklemeyin.
- **Süreler takvim günüdür.** "Ay" ya da "iş günü" olarak yazılmış bildirim sürelerini güne çevirirken ihtiyatlı
  davranın (örneğin 3 ay için 92 gün).
- **Otomatik yenileme bir varsayımdır.** Yazılı bitişi geçmiş sözleşmeyi kod yenilenmiş sayar ve bunu işaretler;
  fiilen yenilendiğini, bedelin ve koşulların değişip değişmediğini doğrulayıp tabloyu güncelleyin.
- **Kapsam dışı:** kira ve bedel artış tarihleri, teminat mektubu vadeleri, sözleşmedeki diğer yükümlülük ve
  cezai şartlar, kanundan doğan özel yenileme ve fesih kuralları.
- **Örnek veri:** Taraflar, tarihler ve tutarlar kurgusaldır.

## Testler

- Fesih bildirim son günü, karar tarihi ve durum sınırları; otomatik yenilemede bitişin ileri taşınması.
- Ay ekleme (ay sonu, 29 Şubat), süresiz sözleşme, bildirim süresi yazılmamış sözleşme.
- Tür ve sorumlu özeti, sıralama, okuma hataları ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
