# Plan Uyum Raporu · Kod Bloğu

> Üretim Planlama › Üretim Planlama Uzmanı · Workers / Workless

Planlanan ile gerçekleşen üretimi **iş emri bazında** karşılaştırır; miktar ve tarih sapmalarını ayırır, sapma
nedenlerini sayar ve plan uyum oranlarını hat ve ürün kırılımıyla verir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Miktar durumu | Gerçekleşen < plan × (1 − tolerans) "Eksik"; > plan × (1 + tolerans) "Fazla"; arası "Tam" (`--tolerans`, %0) |
| Tarih durumu | Gerçekleşen bitiş ≤ planlanan bitiş "Zamanında", sonrası "Geç" |
| Gecikmiş (açık) | Gerçekleşen bitiş boş ve planlanan bitiş durum tarihinden önce |
| Devam ediyor | Gerçekleşen bitiş boş, planlanan bitiş henüz geçmemiş; değerlendirmeye girmez |
| Plan uyumu % | Plana uygun (tam ve zamanında) iş emri ÷ değerlendirilen iş emri × 100 |
| Miktar uyumu % | Σ min(gerçekleşen, plan) ÷ Σ plan × 100 |
| Zamanında bitirme % | Zamanında biten iş emri ÷ değerlendirilen iş emri × 100 |

İş emri durumları: Plana uygun, Eksik, Geç, Eksik ve geç, Fazla üretim, Gecikmiş (açık), Devam ediyor. Fazla
üretim miktar uyumunu artırmaz: bir iş emrindeki fazla, başka iş emrindeki eksiği kapatmaz.

**Örnek:** Plan 3.000, gerçekleşen 2.400, zamanında bitti → "Eksik", gerçekleşme **%80**, eksik **600**. Plan 5.000,
gerçekleşen 5.300 → "Fazla üretim"; miktar uyumunda 5.000 olarak sayılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri, durum tarihi 12.10.2026
python main.py --emirler is_emirleri.xlsx
python main.py --emirler is_emirleri.xlsx --tolerans 2 --tarih 31.10.2026
```

Zorunlu sütunlar: İş Emri No, Planlanan Miktar, Planlanan Bitiş.

## Çıktı

`cikti/plan_uyum_raporu.xlsx`:

- `İş Emirleri`: en sorunludan başlayarak; fark, gerçekleşme oranı, gecikme, sapma nedeni.
- `Özet`: plan uyumu, miktar uyumu, zamanında bitirme, durum dağılımı.
- `Sapma Nedenleri`: neden bazında iş emri sayısı, eksik miktar ve toplam gecikme.
- `Hatlar`, `Ürünler`: miktar uyumu en düşük olandan başlayarak.
- `Uyarılar`.

## Dikkat

- **Sapma nedenini kod atamaz.** Dosyada yazdığınız neden sayılır; sapması olup nedeni boş bırakılan iş emirleri
  ayrıca bildirilir. Neden kodlarını tutarlı yazmazsanız aynı neden birkaç satıra bölünür.
- **Plan, dosyadaki plandır.** Kod planın revize edilip edilmediğini bilmez. Planı gerçekleşmeye göre sonradan
  güncellerseniz uyum yüksek görünür; ilk (dondurulmuş) planı kullanmanız daha doğru bir ölçüm verir.
- **Tolerans politikanızdır.** %0 varsayılanında bir adet eksik bile "Eksik" sayılır. Sipariş toleransınız ya da
  süreç fireniz varsa `--tolerans` ile verin.
- **Gerçekleşen miktarın tanımı sizdedir.** Sağlam (kabul edilen) adet mi, toplam üretilen mi; fire dahil
  yazarsanız eksik üretim gizlenir.
- **Fazla üretim "uygun" sayılmaz.** Plan üstü üretim stok ve kapasite maliyeti doğurabileceği için ayrı durumdur;
  uygun saymak istiyorsanız toleransı artırın.
- **Kapsam dışı:** sıra (sekans) uyumu, kapasite ve malzeme planlaması, sevkiyat ve müşteri teslim performansı,
  kısmi teslimatlar.
- **Örnek veri:** İş emirleri, ürünler ve hatlar kurgusaldır.

## Testler

- Durum ataması: eksik, geç, eksik ve geç, fazla üretim, gecikmiş açık ve devam eden iş emirleri.
- Plan uyumu, miktar uyumu (fazla üretimin kırpılması), zamanında bitirme; tolerans sınırı.
- Sapma nedeni dökümü, nedeni yazılmamış sapmalar, hat ve ürün kırılımı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
