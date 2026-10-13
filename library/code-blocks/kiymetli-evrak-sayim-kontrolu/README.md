# Kıymetli Evrak Sayım Kontrolü · Kod Bloğu

> Şube Bankacılığı › Şube Operasyon Yöneticisi · Workers / Workless

Kıymetli evrak sayım tutanağını kayıtlarla **seri numarası düzeyinde** karşılaştırır: eksik, fazla ve iki kez
sayılmış numaraları aralık olarak listeler. Adetler tutsa bile yanlış seri numarasını yakalar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Seri numarası | Baştaki harfler (ön ek) + sondaki rakamlar; `A 0012501` → ön ek A, numara 12501 |
| Aralık | Seri Başlangıç – Seri Bitiş arasındaki bütün numaralar; bitiş boşsa tek numara |
| Eksik | Kayıtta olup sayımda olmayan seri numaraları |
| Fazla | Sayımda olup kayıtta olmayan seri numaraları |
| Mükerrer sayım | Sayım listesinde birden çok kez geçen seri numaraları (çakışan aralıklar) |
| Sonuç | Fark var (eksik ya da fazla) · Kontrol et (yalnız mükerrer) · Tam |

Karşılaştırma evrak türü ve ön ek bazındadır. Boşluk ve büyük-küçük harf farkı dikkate alınmaz (`a0012501` =
`A 0012501`). Ardışık farklar tek satırda birleştirilir.

**Örnek:** Kayıt `B 0004001 – B 0004040` (40 adet); sayım `B 0004001 – 0004025` ve `B 0004028 – 0004040` (38 adet) →
**eksik: B 0004026 – B 0004027 (2 adet)**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri
python main.py --kayit kayit.xlsx --sayim sayim.xlsx
```

Zorunlu sütunlar iki dosyada da Evrak Türü ve Seri Başlangıç. Aynı tür için birden çok aralık satırı yazılabilir.
Bitişte ön eki tekrar yazmak zorunlu değildir (`A 0012501` – `0012560`). Tek satırda en çok 100.000 numara olabilir.

## Çıktı

`cikti/kiymetli_evrak_sayim_kontrolu.xlsx`:

- `Özet`: evrak türü bazında kayıtlı ve sayılan adet, eksik, fazla, mükerrer sayım ve sonuç.
- `Farklar`: her fark için seri başlangıç ve bitiş, adet; "Açıklama" sütununu siz doldurursunuz.
- `Uyarılar`: eksik ve kayıt dışı evrak, mükerrer sayım ve kayıt, okunamayan satırlar.

## Dikkat

- **Bu bir karşılaştırma aracıdır, sayım tutanağı değildir.** Sayımın çift kişiyle yapılması, tutanağın imzalanması,
  farkın bildirilmesi ve eksik evrakın iptal ya da bloke edilmesi kurumunuzun iç düzenlemelerine göre yürütülür.
- **Kayıt dosyası "olması gereken"dir.** Sayım anına kadar teslim edilmiş, iptal edilmiş ya da başka şubeye
  gönderilmiş evrak kayıttan düşülmemişse eksik görünür. Sayım saatiyle kaydın kesim saatini eşitleyin.
- **Tür adları iki dosyada aynı olmalı.** "Çek Karnesi (10 yaprak)" ile "10'luk Çek Karnesi" ayrı tür sayılır;
  biri tamamen eksik, öbürü tamamen fazla çıkar.
- **Koçan içi yaprak takibi yapılmaz.** Kod yazdığınız seri numarasını bir birim sayar; karne numarasıyla mı yaprak
  numarasıyla mı izlediğiniz iki dosyada aynı olmalıdır.
- **Kapsam dışı:** nakit ve efektif sayımı, kasa ve anahtar kontrolü, müşteriye teslim edilmeyi bekleyen
  kişiselleştirilmiş kartlar, emanet ve rehin kıymetler.
- **Gizlilik:** Seri numaraları hassas operasyonel bilgidir; raporu yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Evrak türleri ve seri numaraları kurgusaldır.

## Testler

- Seri numarası çözümleme (ön ek, baştaki sıfırlar, boşluk ve harf farkı), tek numara, ön eksiz bitiş.
- Eksik, fazla, mükerrer sayım ve mükerrer kayıt; ardışık farkların aralığa birleşmesi, farklı ön eklerin ayrılması.
- Hatalı satırlar (ters aralık, farklı ön ek, çok büyük aralık), sayımda hiç olmayan tür, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
