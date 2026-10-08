# Portföy Müşteri Fırsat Listesi · Kod Bloğu

> Bankacılık › Şube Bankacılığı › Bireysel Portföy Yöneticisi · Workers / Workless

Portföydeki bireysel müşterilerin ürün sahipliği ve hareketlerinden **çapraz satış ve elde tutma fırsatlarını**
kural tabanlı olarak puanlar ve önceliklendirir. Kart, BES, mevduat ve sigorta gibi fırsatlar için gerekçeli bir
arama/görüşme listesi çıkarır. İnternete bağlanmaz.

## Kurallar

Eşikler ve puanlar örnektir; `--ayarlar` ile verilen JSON dosyasıyla değiştirilebilir.

| Fırsat | Koşul | Puan |
|---|---|---|
| Vade yenileme | Vadeli mevduatın vadesi 15 gün içinde | 40 |
| Kredi kartı | Kartı yok, gelir ≥ 25.000 TL, gecikmesi yok | 30 (+10 maaş müşterisi) |
| Vadeli mevduat / yatırım fonu | 3 aylık ort. vadesiz bakiye ≥ 150.000 TL, vadeli ve fonu yok | 30 (+10, ≥ 500.000 TL) |
| BES | BES'i yok, 18-56 yaş, gelir ≥ 20.000 TL | 25 (+10 maaş müşterisi) |
| Kasko | Taşıt kredisi var, kaskosu yok | 25 |
| Konut sigortası | Konut kredisi var, konut sigortası yok | 20 |
| Reaktivasyon | 90 günden uzun süredir işlem yok | 20 |
| Hayat sigortası | Kredisi var, hayat sigortası yok, 25-60 yaş | 15 |
| Mobil bankacılık | Kullanmıyor | 15 |
| Otomatik ödeme talimatı | Maaş müşterisi, talimatı yok | 15 |
| İhtiyaç görüşmesi | Kredisi 60 gün içinde kapanıyor | 10 |

**Sorumlu satış:**
- Gecikmesi olan veya kart limit kullanımı %90 ve üzeri olan müşteriye kart, kredi ve ihtiyaç görüşmesi önerilmez;
  not düşülür.
- Sigortalar krediye bağlı zorunlu ürün olarak sunulamaz; gerekçeye not yazılır.

**İletişim izni:** Pazarlama (İYS) izni olmayan müşterilerde kanal "yalnız şube görüşmesi" olarak işaretlenir;
`--izinsiz-haric` ile bu müşteriler listeden çıkarılır.

## Kullanım

```bash
pip install -r requirements.txt
python main.py                                                   # örnek: 12 müşteri, rapor tarihi 08.10.2026
python main.py --girdi portfoy.xlsx --tarih 08.10.2026 --ilk 2
python main.py --girdi portfoy.xlsx --ayarlar kurallar.json --izinsiz-haric
```

**Girdi sütunları:**
- Müşteri No, Ad Soyad, Yaş, Maaş Müşterisi (E/H), Aylık Gelir.
- Vadesiz Ort. Bakiye (3 Ay), Vadeli Mevduat, Vadeli Vade Tarihi, Yatırım Fonu.
- Kredi Kartı, Kart Limit Kullanımı (%), BES, Hayat Sigortası, Konut Kredisi, Konut Sigortası, Taşıt Kredisi,
  Kasko, İhtiyaç Kredisi, Kredi Bitiş Tarihi.
- Mobil Bankacılık, Otomatik Ödeme Talimatı (sayı), Son İşlem Tarihi, Gecikme, Pazarlama İzni.

Eksik sütun varsa ilgili kural uygulanmaz ve rapora not düşülür.

**Ayar dosyası örneği:** `{"min_gelir_kart": 30000, "vadesiz_esik": 250000, "puan": {"bes": 35}}`

## Çıktı

`portfoy_firsat_listesi.xlsx`:
- `Fırsat Listesi`: puana göre sıralı; gerekçe, kanal, "Görüşme Tarihi" ve "Sonuç" sütunları.
- `Müşteri Özeti`: ürün sayısı, ilk 3 fırsatın puanı.
- `Ürün Penetrasyonu`: ürün sahipliği oranları, grafikli.
- `Kurallar`.

## Dikkat

- **Ön liste niteliği:** Liste bir **görüşme önceliklendirmesidir**, satış kararı değildir. Ürün uygunluğunu
  bankanızın politikasına göre değerlendirin.
- **Yatırım ürünleri:** Müşterinin risk profili ve uygunluk testi ayrıca değerlendirilmelidir.
- **Ticari elektronik ileti:** SMS, e-posta ve arama için müşterinin onayı (İYS) gerekir.
- **Kişisel veri ve bankacılık sırrı:** Portföy listesi kişisel veri ve bankacılık sırrı içerir; dosyaları
  yetkisiz kişilerle paylaşmayın.
- **Örnek veri:** Örnek müşteriler kurgusaldır.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçlar kontrol edilmeden işlem yapılmamalıdır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
