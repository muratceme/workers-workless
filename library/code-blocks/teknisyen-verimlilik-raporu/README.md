# Teknisyen Verimlilik Raporu · Kod Bloğu

> Satış Sonrası Hizmetler (Servis) › Satış Sonrası Hizmetler Müdürü · Workers / Workless

İş emirlerinden teknisyen bazında **satılan saat** (faturalanan / standart süre), **çalışılan saat** (işe kayıtlı süre) ve
puantajdaki **mevcut saati** toplayıp atölyenin üç temel oranını hesaplar: verimlilik, kullanım ve üretkenlik. İnternete ve DMS'e
bağlanmaz.

## Nasıl hesaplar?

| Oran | Formül | Ne gösterir |
|---|---|---|
| Verimlilik % | Satılan saat ÷ çalışılan saat × 100 | İşin standart süreye göre ne kadar hızlı yapıldığı |
| Kullanım % | Çalışılan saat ÷ mevcut saat × 100 | İşyerindeki sürenin ne kadarının işe kayıtlı geçtiği |
| Üretkenlik % | Satılan saat ÷ mevcut saat × 100 (= verimlilik × kullanım ÷ 100) | Mevcut her saatin kaç saat satışa dönüştüğü |
| Boş saat | Mevcut − çalışılan | İşe kayıtlı olmayan süre |

Atölye satırında kullanım ve üretkenlik yalnız puantajı olan teknisyenlerden hesaplanır. İş türü kırılımı (Müşteri / Garanti /
İç) yalnız satılan, çalışılan ve verimliliği gösterir.

**Uyarılar:** bir teknisyenin bir günde iş emirlerine yazılan süresi puantajdaki mevcut saati aşıyor (Yüksek) · puantajı olmayan
teknisyen (Orta) · verimlilik `--verim-alt` (%85) ya da kullanım `--kullanim-alt` (%80) altında · verimlilik %150'nin üstünde
(süre kaydı eksik olabilir) · puantajı olup iş emri olmayan teknisyen (Bilgi).

**Örnek:** 9 satılan saat, 8 çalışılan saat, 15 mevcut saat → verimlilik **%112,50**, kullanım **%53,33**, üretkenlik **%60**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                            # örnek veri (eylül 2026, 6 teknisyen)
python main.py --is-emirleri is_emirleri.xlsx --puantaj puantaj.xlsx
python main.py --is-emirleri is_emirleri.xlsx --puantaj puantaj.xlsx --verim-alt 90 --kullanim-alt 85
python main.py --is-emirleri is_emirleri.xlsx                              # yalnız verimlilik
```

## Çıktı

`cikti/teknisyen_verimlilik.xlsx`: `Teknisyen` (üretkenliğe göre sıralı; eşik altı sarı, atölye toplamıyla), `İş Türü`, `Uyarılar`.

## Dikkat

- **Oranlar süre kaydının kalitesine bağlıdır.** İşe başlama / bitirme (clock-on / off) kayıtları eksik ya da toplu girilmişse
  verimlilik ve kullanım yanıltıcı olur.
- **Eşikler bilgi amaçlıdır.** Düşük oran; iş dağılımı, parça bekleme, eğitim, tekrar iş ya da randevu planından kaynaklanabilir.
  Rapor tek başına performans değerlendirmesi ya da ücret / disiplin kararı için kullanılmamalıdır; çalışan verisi KVKK
  kapsamındadır.
- **Garanti işlerinde satılan saat** markanın ödediği standart süredir; müşteri işleriyle ayrı değerlendirin.
- **Örnek veri:** Teknisyenler ve iş emirleri kurgusaldır.

## Testler

- Teknisyen ve atölye bazında satılan / çalışılan / mevcut saat, verimlilik, kullanım, üretkenlik, boş saat, iş türü kırılımı ve sıralama.
- Günlük süre aşımı, puantaj eksikliği ve eşik uyarıları; okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
