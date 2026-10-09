# Personel Maliyet Bütçesi · Kod Bloğu

> İnsan Kaynakları › İnsan Kaynakları Müdürü · Workers / Workless

Kadro planı, ücretler ve yan haklardan aylık personel maliyet bütçesi üretir. Bütçeyi departman ve ay bazında
gerçekleşmeyle karşılaştırır. İnternete bağlanmaz.

## Nasıl hesaplar?

- **Kadro:** Her satır bir kişi veya planlanan kadrodur.
  - Başlangıç ayı boşsa kişi mevcut kadrodadır. Doluysa yıl içinde başlayan yeni kadrodur.
  - Bitiş ayı planlanan ayrılışı gösterir.
  - Giriş ve çıkış ayları tam ay sayılır.
- **Ücret artışı:** `--zam 1=25 7=10` → Ocak'tan itibaren %25, Temmuz'dan itibaren ayrıca %10 (birikimli).
  Mevcut kadroya ve artış ayından önce başlayan yeni kadroya uygulanır.
- **Yan haklar:**
  - TL veya brüt ücretin %'si (ör. ikramiye %100).
  - Kapsam: `Herkes`, `Departman=Üretim`, `Departman!=Satış`, `Pozisyon~Müdür`.
  - SGK'ya tabi mi, hangi aylarda ödendiği (`*` her ay, `6,12`).
- **İşveren maliyeti:** brüt ücret + yan haklar + SGK işveren payı + işsizlik sigortası işveren payı.
  - Prime esas kazanç = brüt + SGK'ya tabi yan haklar. SGK tavanıyla sınırlıdır.
  - Teşvik puanı (yok / imalat / diger) SGK işveren payından düşülür.
  - Oranlar ve tavan ortak `tr_parametreler.json` dosyasından yıl bazında okunur. Bütçe yılının değerleri yoksa
    önceki yılınkiler kullanılır ve uyarı verilir.
- **Gerçekleşme:** Departman × ay tutarı bütçeyle karşılaştırılır. `--esik` (varsayılan %5) aşan sapma Orta,
  3 katını aşan sapma Yüksek önemdedir. Bütçesi olmayan gerçekleşme ve gerçekleşmesi olmayan bütçe ayrıca listelenir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 14 kadro, 2026, Temmuz'da %10 artış, 9 ay gerçekleşme
python main.py --kadro kadro.xlsx --yan-haklar yan_haklar.csv --yil 2027 --zam 1=25 7=10
python main.py --kadro kadro.xlsx --yan-haklar y.csv --yil 2026 --gerceklesen gerceklesen.xlsx --esik 3
```

| Dosya | Sütunlar |
|---|---|
| Kadro | Kadro No, Pozisyon, Departman, Brüt Ücret (TL), Başlangıç Ayı, Bitiş Ayı, Teşvik (yok / imalat / diger) |
| Yan haklar | Kalem, Tür (TL / %), Tutar, Kapsam, SGK'ya Tabi (Evet / Hayır), Aylar |
| Gerçekleşen (isteğe bağlı) | Departman, Ay (1–12 veya ay adı), Tutar (TL) |

Ay sütunlarına `4`, `Nisan`, `2027-04` veya `04.2027` yazılabilir.

## Çıktı

`personel_maliyet_butcesi.xlsx`:
- `Aylık Bütçe`: departman × ay, yıllık toplam, kişi sayısı ve yığılmış sütun grafiği.
- `Maliyet Bileşenleri`: brüt ücret, her yan hak, SGK ve işsizlik payı; ay bazında ve toplam içindeki payı.
- `Kişi Bazında`: yıllık brüt, yan haklar, SGK + işsizlik, işveren maliyeti, aylık ortalama, SGK tavanı işareti.
- `Gerçekleşme`: bütçe, gerçekleşen, sapma, sapma %; boş "Açıklama" sütunu.
- `Varsayımlar`: kullanılan oranlar, tavan, teşvik puanları, artışlar, yan haklar ve kapsam dışı kalemler.
- `Uyarılar`.

## Dikkat

- **Kapsam:** Kıdem tazminatı karşılığı, ihbar ve izin karşılığı, fazla mesai, damga vergisi ve kıst ay hesabı dahil
  değildir. SGK tavanını aşan ikramiye tutarının sonraki aylara devri hesaplanmaz.
- **SGK'ya tabi olma:** Yan hakların SGK ve gelir vergisi karşısındaki durumu ve istisna sınırları (yemek, yol, özel
  sağlık sigortası, araç tahsisi vb.) mevzuata göre belirlenir. Örnek dosyadaki Evet / Hayır değerleri örnektir;
  bordro uygulamanız ve mali müşavirinizle doğrulayın.
- **Teşvikler:** Yalnız genel işveren hissesi indirimi (5 / 2 puan) modellenir. Diğer teşvikler ayrıca
  değerlendirilmelidir; örnekler: asgari ücret desteği, istihdam teşvikleri.
- **Parametreler:** Yeni yılın SGK oranları, tavanı ve asgari ücreti açıklandığında `tr_parametreler.json` kaynağıyla
  birlikte güncellenmelidir.
- **Örnek veri:** Örnek kadro, ücretler ve gerçekleşme verisi kurgusaldır.

## Testler

Şunlar test edilir:
- Aylık işveren maliyeti: SGK teşviki, işsizlik payı, SGK tavanı.
- Ücret artışı (mevcut ve yeni kadro), ikramiye ayları, giriş-çıkış ayları, yan hak kapsamı.
- Gerçekleşme sapmaları ve eksik gerçekleşme.
- Parametre yedeği, asgari ücret uyarısı, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
