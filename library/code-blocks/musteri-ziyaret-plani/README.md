# Müşteri Ziyaret Planı · Kod Bloğu

> Satış › Satış Temsilcisi · Workers / Workless

Müşteri listesi, konum ve ziyaret sıklığından **haftalık ziyaret planı** çıkarır. Bu hafta ziyareti gelen müşterileri
bulur, birbirine yakın olanları aynı güne toplar ve her gün için **rota sırası** ile tahmini varış saatlerini verir.
İnternete bağlanmaz; harita servisi kullanmaz.

## Nasıl hesaplar?

1. **Ziyareti gelenler:** Sonraki ziyaret = son ziyaret + sıklık. Bu tarih haftanın son plan gününe kadar doluyorsa
   müşteri plana girer. Hiç ziyaret edilmemiş müşteri de girer.

   | Sıklık yazımı | Gün |
   |---|---|
   | Haftalık | 7 |
   | 2 haftada bir | 14 |
   | Aylık | 28 |
   | 3 ayda bir | 84 |
   | 10 gün | 10 |

2. **Günlere dağıtım (süpürme yöntemi):** Müşteriler başlangıç noktasına göre açı sırasına dizilir ve ardışık
   gruplara bölünür. Böylece her gün tek bir yöne gidilir.
   - Grup büyüklüğü günlük ziyaret süresine göre dengelenir.
   - "Uygun Günler" kısıtı olan müşteri, izinli günlerden rotayı en az uzatanına eklenir.
3. **Rota sırası:** En yakın komşu yöntemiyle kurulur, 2-opt ile iyileştirilir. Tur başlangıç noktasından çıkar ve
   oraya döner.
4. **Mesafe ve süre:**
   - Mesafe = kuş uçuşu mesafe × `--yol-katsayisi` (1,3).
   - Süre = mesafe / `--hiz` (30 km/s) + ziyaret süreleri.
5. **Kapasite:** Günün toplam süresi `--gunluk-dk` (480) değerini aşarsa müşteri çıkarılır ve "Sığmayanlar"a yazılır.
   Önce önceliği en düşük (C → B → A), eşitlikte gecikmesi en az olan çıkar.

| Uyarı | Önem |
|---|---|
| A öncelikli müşteri plana sığmadı | Yüksek |
| B / C müşteri plana sığmadı; konumu olmayan müşteri | Orta |
| Ziyaret, sıklık süresinden fazla gecikmiş | Orta |
| Sıklık okunamadı (aylık kabul edilir) | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek: 48 müşteri, 12.10.2026 haftası
python main.py --musteriler musteriler.xlsx --hafta 12.10.2026 --baslangic 40.99,29.12
python main.py --musteriler m.xlsx --hafta 12.10.2026 --baslangic 40.99,29.12 --gun-sayisi 6 --gunluk-dk 420 --hiz 40
```

| Sütun | Açıklama |
|---|---|
| Müşteri No, Müşteri, İlçe | Tanım |
| Enlem, Boylam | Ondalık derece (ör. 40,98512 ve 29,04233). Harita uygulamasında noktaya sağ tıklayarak alınabilir. |
| Ziyaret Sıklığı | Yukarıdaki yazımlardan biri |
| Ziyaret Süresi (dk) | Boşsa `--ziyaret-dk` (30) |
| Öncelik | A / B / C |
| Son Ziyaret | Boşsa hiç ziyaret edilmemiş sayılır |
| Uygun Günler | İsteğe bağlı: Pzt, Sal, Çar, Per, Cum, Cmt |

`--baslangic` kendi verinizle çalışırken zorunludur (ofis, depo veya ev konumu).

## Çıktı

`ziyaret_plani.xlsx`:
- `Haftalık Plan`: gün, sıra, müşteri, tahmini varış saati, ziyaret süresi, önceki noktadan mesafe, gecikme, harita
  bağlantısı; boş "Ziyaret Notu" sütunu.
- `Gün Özeti`: ziyaret sayısı, toplam yol, yol ve ziyaret süresi, kapasite kullanımı, tahmini bitiş saati.
- `Sığmayanlar`: plana giremeyen müşteriler ve nedeni.
- `Bu Hafta Gerekmeyenler`: sonraki ziyaret tarihleriyle.
- `Uyarılar`.

## Dikkat

- **Mesafe ve süreler tahmindir.** Gerçek yol mesafesi, trafik, köprü ve otopark süreleri hesaba katılmaz. Yoğun
  trafikli bölgelerde hızı düşürün veya yol katsayısını artırın.
- **Rota yaklaşık en iyidir.** 2-opt iyi sonuç verir ama en kısa turu garanti etmez.
- **Randevu saatleri** dikkate alınmaz. Randevulu ziyaretleri "Uygun Günler" ile güne sabitleyip sırayı elle
  düzenleyin.
- **Kişisel veri:** Bireysel müşterilerin adres ve konum bilgisi kişisel veridir. Harita bağlantısı yalnız tıklandığında
  tarayıcıda açılır; paket konumları hiçbir yere göndermez.
- **Örnek veri:** Örnek müşteriler ve konumlar kurgusaldır.

## Testler

Şunlar test edilir:
- Haversine mesafesi (1 enlem derecesi ≈ 111,2 km) ve doğru üzerindeki noktalarda 2-opt turu.
- Sıklık ve gün yazımları.
- Vade kuralı, uygun gün kısıtı, plan günleri dışındaki kısıt, konumsuz müşteri.
- Süre aşımında en düşük öncelikli müşterinin çıkarılması.
- Örnek veride her müşterinin tek kez planlanması, günlük süre sınırı, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
