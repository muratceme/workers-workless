# Ürün Gamı (Delist) Önerisi · Kod Bloğu

> Satın Alma ve Kategori Yönetimi › Kategori Yöneticisi · Workers / Workless

Her ürünü **kendi kategorisi içinde** ciro, brüt kâr ve raf verimliliğine göre değerlendirir; listeden çıkarılabilecek
ürünleri **gerekçesiyle** önerir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Ölçüt | Kural |
|---|---|
| Ciro kuyruğu | Ürünler ciroya göre sıralandığında, kendisinden önceki ürünler kategori cirosunun %(100 − `--kuyruk`) kadarını zaten oluşturuyorsa. Varsayılan `--kuyruk 10`: cironun son %10'unu oluşturan ürünler |
| Kâr kuyruğu | Aynı kural brüt kâra göre; brüt kârı sıfır ya da eksi olan ürün her zaman kuyruktadır |
| Raf verimliliği | Kâr payı ÷ raf payı × 100 (yalnız raf alanı yazılmış ürünler arasında); `--raf` (varsayılan 50) altı düşük |

| Öneri | Ne zaman |
|---|---|
| Delist adayı | Değerlendirilebilen ölçütlerin hepsi olumsuz (raf alanı yoksa iki ölçüt, varsa üç) |
| İzle | Üç ölçütten ikisi olumsuz |
| Devam | Diğerleri |
| Stratejik | "Stratejik" sütununda Evet / x yazan ürün; öneri verilmez |
| Yeni ürün | Listeleme tarihi `--yeni` (varsayılan 90) günden yeni; öneri verilmez |

Brüt kâr yazılmamışsa ciro − maliyet kullanılır.

**Örnek:** Şampuan kategorisinin cirosu 704.500. S01–S04 birlikte 643.000 (%91) ediyor; S05 (31.500) bu yüzden
ciro kuyruğunda. Kâr kuyruğunda da; raf payı %17,39, kâr payı %2,49 → raf verimliliği 14,33 → **Delist adayı**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                           # örnek veri (tarih 30.09.2026)
python main.py --urunler kategori_urun_performansi.xlsx --tarih 30.09.2026
python main.py --urunler kategori_urun_performansi.xlsx --kuyruk 15 --raf 40 --yeni 180
```

Zorunlu sütunlar: Ürün Kodu, Ciro ve Brüt Kâr ya da Maliyet. Raf alanı santimetre, metre ya da yüz (facing) sayısı
olabilir; aynı kategoride tek birim kullanın.

## Çıktı

`cikti/urun_gami_delist_onerisi.xlsx`:

- `Delist Adayları`: delist adayı ve izlenecek ürünler, gerekçeleriyle.
- `Tüm Ürünler`: ciro, kâr ve raf payları, marj, raf verimliliği ve öneri.
- `Kategoriler`: adayların kategori cirosu, kârı ve rafındaki payı (çıkarılırsa ne kaybedilir, ne kadar raf açılır).
- `Uyarılar`.

## Dikkat

- **Öneri karar değildir.** Ürün çıkarma kararında ikame (satışın başka ürüne kayıp kaymayacağı), müşteri sadakati,
  tamamlayıcı ürünler, tedarikçi anlaşmaları (listeleme bedeli, ciro primi, iade koşulları) ve stoktaki ürünlerin
  eritilmesi ayrıca değerlendirilmelidir. Kod bunları bilmez.
- **Dönem ve kapsam aynı olmalı.** Tüm ürünlerin verisi aynı tarih aralığına ve aynı mağaza kapsamına ait olmalıdır;
  dönem içinde stoksuz kalan ya da sonradan listelenen ürünlerin satışı düşük görünür.
- **Eşikler politika varsayılanıdır.** %10 kuyruk, 50 raf verimliliği ve 90 gün yeni ürün süresi genel kabul görmüş
  değerler değildir; kategori yapınıza göre değiştirin.
- **Brüt kâr tanımı:** Ciro primi, listeleme bedeli gibi arka marj gelirleri brüt kâra dahil edilmezse bazı ürünler
  olduğundan zayıf görünür. Tanımı kategoriler arasında tutarlı kullanın.
- **Örnek veri:** Ürünler ve sayılar kurgusaldır.

## Testler

- Ciro ve kâr kuyruğu, zarar eden ve satışı olmayan ürün, raf verimliliği, ölçüt sayısına göre öneri.
- Stratejik ve yeni ürün istisnası, kategori özeti, eşikler, raf alanı olmadan değerlendirme, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
