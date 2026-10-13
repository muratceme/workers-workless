# Ay Sonu Kapanış Kontrol Listesi · Kod Bloğu

> Muhasebe › Genel Muhasebe Uzmanı · Workers / Workless

Kapanış adımlarınızı **hedef gün, sorumlu, önkoşul ve durumla** izler: hangi adım gecikti, hangisi bugün bitmeli,
hangisi başka bir adımı bekliyor, şu an hangileri yapılabilir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Hedef tarih | "Gün" sütunu N: `0` dönemin son iş günü, `3` ondan sonraki 3. iş günü, `-2` önceki 2. iş günü |
| İş günü | Hafta içi ve `--tatiller` dosyasında olmayan gün |
| Gecikti | Tamamlanmamış ve hedefi durum tarihinden önce; gecikme takvim günüdür |
| Gecikmeli tamamlandı | Tamamlanma tarihi hedeften sonra |
| Engelleyen adımlar | Önkoşullardan henüz kapanmamış olanlar ("Uygulanmaz" engel sayılmaz) |
| Tamamlanma oranı | Tamamlanan ÷ (tüm adımlar − uygulanmaz) |

**Örnek:** Dönem 09.2026 (son gün 30.09 Çarşamba). Gün `1` → 01.10, `3` → 05.10 (hafta sonu atlanır), `-2` → 28.09.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri, dönem 09.2026, durum tarihi 06.10.2026
python main.py --adimlar adimlar.xlsx --donem 09.2026
python main.py --adimlar adimlar.xlsx --donem 09.2026 --tarih 06.10.2026 --tatiller tatiller.csv
```

Zorunlu sütunlar: Adım, Gün. Diğerleri: Sıra, Kategori, Sorumlu, Önkoşul (virgülle sıra numaraları), Durum
(Tamamlandı / Devam ediyor / Başlamadı / Uygulanmaz), Tamamlanma Tarihi, Not. `--donem` verilmezse durum tarihinden
önceki ay alınır.

## Çıktı

`cikti/kapanis_kontrol_listesi.xlsx`:

- `Kontrol Listesi`: her adımın hedef tarihi, izleme durumu, gecikme günü ve engelleyen adımları.
- `Kategori Özeti`, `Sorumlu Özeti`: adım, tamamlanan, geciken ve tamamlanma yüzdesi.
- `Uyarılar`: geciken adımlar, önkoşulu açıkken kapatılanlar, sorumlusuz adımlar, okunamayan satırlar.

## Dikkat

- **Adımları siz belirlersiniz.** Kod, şirketinizde hangi kapanış kayıtlarının gerektiğini bilmez; örnek liste
  genel bir başlangıçtır, kendi hesap planınıza ve süreçlerinize göre düzenleyin.
- **Kod kayıt atmaz ve tutar kontrol etmez.** Bir adımın "Tamamlandı" yazılması, mutabakatın tuttuğunu göstermez.
- **Resmî tatilleri siz verirsiniz.** Dosya verilmezse yalnız hafta sonları atlanır; yarım günler tam iş günü sayılır.
- **Yasal süreler ayrı izlenir.** Beyanname ve bildirim son günleri bu listenin konusu değildir.
- **Örnek veri:** Kişiler ve tarihler kurgusaldır.

## Testler

- İş günü hesabı (hafta sonu, tatil, negatif gün, ay sonu hafta sonuna denk gelen dönem).
- İzleme durumları, önkoşul ve engel mantığı, özetler ve tamamlanma oranı.
- Okuma hataları, tanınmayan durum, komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
