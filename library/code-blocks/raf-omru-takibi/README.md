# Raf Ömrü Takibi · Kod Bloğu

> Kalite Kontrol ve Laboratuvar › Kalite Kontrol Teknisyeni · Workers / Workless

Parti stoklarının **son tüketim tarihine kalan süresini** ve **kalan ömür yüzdesini** hesaplar; her partinin hangi
müşterilerin **kabul kriterini** karşıladığını gösterir ve planlanan sevkiyatları sevk günündeki kalan ömre göre kontrol
eder. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Son tüketim tarihi | STT / TETT sütunu; yoksa üretim tarihi + raf ömrü (gün) |
| Kalan ömür | Son tüketim tarihi − değerlendirme günü (gün) |
| Kalan ömür % | Kalan gün ÷ (son tüketim − üretim) × 100 |
| Müşteri kriteri | Müşteri dosyasında en az kalan ömür % ya da en az kalan gün; ürüne özel kural genel kuraldan önce gelir; müşteri kuralı yoksa `--kabul` (varsayılan %66) |

| Stok durumu | Kural |
|---|---|
| Süresi dolmuş | Son tüketim tarihi geçmiş |
| Kabul altı | Kalan ömür `--kabul` yüzdesinin altında |
| Yaklaşan | Kalan gün `--yakin` (varsayılan 30) ve altı |
| Uygun | Diğerleri |

Sevkiyat kontrolünde her satır **sevk tarihindeki** kalan ömürle, o müşterinin kuralına göre değerlendirilir.

**Örnek:** Zincir Market A yoğurtta en az 20 gün kalan ömür istiyor. 05.10.2026 üretimli, 30 gün raf ömürlü partinin
16.10.2026'da 19 günü kalıyor → **uygun değil**. Aynı parti %50 kriterli Toptancı C için uygun (%63).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                       # örnek veri (15.10.2026)
python main.py --stok parti_stoklari.xlsx --musteriler musteri_kabul_kriterleri.xlsx --sevkiyat planlanan_sevkiyat.xlsx
python main.py --stok parti_stoklari.xlsx --kabul 50 --yakin 15
```

- **Stok:** Ürün Kodu, Parti No, Son Tüketim Tarihi (ya da Üretim Tarihi + Raf Ömrü); isteğe bağlı Ürün Adı, Miktar, Depo.
- **Müşteriler (isteğe bağlı):** Müşteri, Ürün Kodu (boşsa tüm ürünler), Kabul Kriteri % ve / veya En Az Kalan Gün.
- **Sevkiyat (isteğe bağlı):** Sevk Tarihi, Müşteri, Ürün Kodu, Parti No, Miktar.

## Çıktı

`cikti/raf_omru_takibi.xlsx`:

- `Partiler`: son tüketim tarihine göre sıralı; kalan gün, kalan ömür %, durum, kriterini karşıladığı müşteriler.
- `Sevkiyat Kontrolü`: her sevk satırının sonucu ve gerekçesi.
- `Uyarılar`.

## Dikkat

- **Kabul kriterleri müşteri sözleşmesinden alınmalıdır.** Zincir marketler ve ihracat müşterileri farklı kalan ömür
  şartları uygular; varsayılan %66 yalnız bir başlangıç değeridir.
- **Son tüketim ve tavsiye edilen tüketim tarihi** ürünün etiketindeki tarihtir; raf ömrü çalışması ve etiketleme
  mevzuatı ayrıca yürütülür. Süresi dolmuş ürünün sevki ve imhası mevzuata ve kurum prosedürüne göre yapılır.
- **Depolama koşulu** (soğuk zincir kırılması) raf ömrünü kısaltabilir; kod yalnız tarihe bakar.
- **Örnek veri:** Ürünler, partiler ve müşteriler kurgusaldır.

## Testler

- Son tüketim tarihinin raf ömründen hesaplanması, kalan gün ve yüzde, stok durumu, partiye uygun müşteriler.
- Ürüne özel / genel müşteri kuralı, varsayılan kriter, sevkiyat kontrolü, stokta olmayan parti, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
