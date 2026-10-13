# Muallak Hasar Raporu · Kod Bloğu

> Hasar Yönetimi › Hasar Müdürü · Workers / Workless

Açık hasar dosyalarının **branş, yaş ve tutar dağılımını** çıkarır; kapanan dosyalarda ödenen tutarı ilk rezervle
karşılaştırarak **rezerv yeterliliği göstergelerini** verir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Açık dosya | Kapanış tarihi boş olan dosya |
| Yaş (gün) | Durum tarihi − ihbar tarihi; dilimler 0–90, 91–180, 181–365, 365+ |
| Tutar dilimi | Güncel muallağa göre; `--dilimler` (50000;250000;1000000), alt sınır dahil |
| Tahmini nihai hasar | Güncel muallak + bugüne kadar ödenen (açık dosya) |
| Rezerv gelişimi % | (Tahmini nihai hasar − ilk muallak) ÷ ilk muallak × 100; açık dosyaların toplamı üzerinden |
| Gerçekleşme oranı % | Ödenen ÷ ilk muallak × 100; kapanan dosyaların toplamı üzerinden |
| Yetersiz rezervli dosya | Kapanan dosyada ödenen > ilk muallak × (1 + `--sapma` %20) |
| Fazla rezervli dosya | Kapanan dosyada ödenen < ilk muallak × (1 − `--sapma` %20) |

İlk muallağı sıfır olan dosyalar oranlara girmez, ayrıca sayılır. Bir branşta en az `--min` (3) dosya varsa ve oran
sapma bandının dışındaysa uyarı üretilir.

**Örnek:** Bir branşta kapanan üç dosyanın ilk muallağı 380.000, ödenen toplamı 531.000 → gerçekleşme oranı
**%139,74**: ilk rezervler ortalamada yetersiz kalmış.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                           # örnek veri, durum tarihi 12.10.2026
python main.py --dosyalar hasar_dosyalari.xlsx
python main.py --dosyalar hasar_dosyalari.xlsx --sapma 15 --min 10 --dilimler 100000;500000 --tarih 31.12.2026
```

Zorunlu sütunlar: Dosya No, İhbar Tarihi. Açık ve dönem içinde kapanan dosyaları aynı dosyada verin; kapanış tarihi
dolu olanlar kapalı sayılır.

## Çıktı

`cikti/muallak_hasar_raporu.xlsx`:

- `Özet`: açık dosya ve muallak toplamı, yaş ve tutar dağılımı, kapanan dosyalarda gerçekleşme oranı.
- `Branşlar (Açık)`: dosya, muallak, pay, ortalama, rezerv gelişimi ve yaş dilimlerine göre tutar.
- `Rezerv Yeterliliği`: branş bazında kapanan dosyalarda ilk muallak, ödenen ve gerçekleşme oranı.
- `Açık Dosyalar`, `Kapanan Dosyalar`, `Uyarılar`.

## Dikkat

- **Aktüeryal karşılık hesabı değildir.** Kod hasar gelişim üçgeni kurmaz; gerçekleşmiş ama rapor edilmemiş hasar
  (IBNR) karşılığı, muallak yeterlilik hesabı ve mevzuattaki karşılık testleri kapsam dışıdır. Buradaki oranlar
  yönetim göstergesidir.
- **Gerçekleşme oranı yalnız kapanan dosyalara bakar.** Kolay dosyalar erken kapanır, büyük ve ihtilaflı dosyalar
  açık kalır; bu yüzden kapananlardaki oran açık portföy için iyimser olabilir. Açık dosyalardaki rezerv gelişimiyle
  birlikte okuyun.
- **"İlk muallak" tanımı sizindir.** Dosya açılışında otomatik atanan tutar mı, ilk eksper tahmini mi; bütün
  dosyalarda aynı tanımı kullanın.
- **Ödemesiz kapanan dosyalar oranı düşürür.** Reddedilen ya da vazgeçilen hasarlar "fazla rezervli" görünür;
  bunlar ayrı sütunda sayılır.
- **Ödenen tutarın kapsamı sizdedir.** Rücu ve sovtaj tahsilatları düşülmüş mü, masraflar dahil mi; kod ayırmaz.
- **Az dosyalı branşta oran yanıltır.** Tek büyük dosya oranı belirler; `--min` ile uyarı eşiğini ayarlayın.
- **Örnek veri:** Dosyalar ve tutarlar kurgusaldır.

## Testler

- Yaş ve tutar dilimi sınırları, açık dosyalarda tahmini nihai hasar ve rezerv gelişimi.
- Kapanan dosyalarda sapma bandı sınırları (tam sınırdaki dosya uyumlu sayılır), ödemesiz ve ilk rezervi olmayan dosyalar.
- Branş uyarılarında en az dosya koşulu, kapalı dosyada kalan muallak, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
