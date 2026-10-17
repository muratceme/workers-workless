# Kategori Performans Raporu · Kod Bloğu

> Satın Alma ve Kategori Yönetimi › Kategori Uzmanı · Gerisi Bende

Ürün, marka ve kategori bazında **satış, marj, stok devir hızı ve payı** raporlar; düşük performanslı ürünleri
işaretler. İnternete bağlanmaz; pazar verisi indirmez.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Brüt kâr, marj % | Net satış − SMM; brüt kâr ÷ net satış × 100 |
| Stok gün | Stok (maliyetle) ÷ SMM × dönem gün sayısı |
| Stok devir hızı | SMM ÷ stok (dönem içi) |
| GMROI | Brüt kâr ÷ stok |
| Büyüme % | (Net satış − önceki dönem satış) ÷ önceki dönem satış × 100 |
| Kategori içi pay % | Ürünün veya markanın cirosu ÷ kategorinin cirosu × 100 |
| Pazar payı % | Kategori cirosu ÷ verdiğiniz pazar cirosu × 100 |
| ABC | Ürünler ciroya göre sıralanır; birikimli pay ≤ %80 A, ≤ %95 B, kalanı C |

| İşaret | Koşul |
|---|---|
| Düşük marj | Marj < `--marj-esik` (varsayılan %15) |
| Yavaş dönen stok | Stok gün > `--stok-gun` (varsayılan 90) |
| Satış düşüyor | Büyüme ≤ −`--dusus-esik` (varsayılan %20) |
| Satış yok | Net satış 0 |

**Performans:** "Satış yok" veya iki ve üzeri işaret → Düşük; tek işaret → İzlenecek; işaret yok → İyi.

**Örnek:** net satış 60.000, SMM 54.000, stok 216.000, önceki dönem 90.000 (30 günlük veri) → marj %10, stok gün
216.000 ÷ 54.000 × 30 = 120, büyüme −%33,33 → üç işaret → Düşük.

Grup satırlarında (marka, kategori) oranlar toplamlardan hesaplanır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (30 günlük)
python main.py --satislar urun_satislari.xlsx --pazar pazar_verisi.xlsx --donem-gun 90 --marj-esik 20 --stok-gun 60
```

| Dosya | Sütunlar |
|---|---|
| Ürün satışları | Ürün, Net Satış, SMM (zorunlu); Kategori, Marka, Satış Adedi, Stok (maliyet), Önceki Dönem Satış |
| Pazar verisi (isteğe bağlı) | Kategori, Pazar Cirosu |

## Çıktı

`cikti/kategori_performans.xlsx`:

- `Kategori Özeti`: kategori bazında satış, marj, stok gün, GMROI, büyüme, ciro payı ve pazar payı.
- `Marka Bazında`: kategori içinde marka göstergeleri ve payı.
- `Ürün Bazında`: performans, ABC sınıfı, göstergeler ve işaretler; boş "Aksiyon" sütunu.
- `Uyarılar`: okunamayan değerler, maliyeti satışından büyük ürünler.

## Dikkat

- **Eşikler kategoriye göre değişir.** Taze ürün ile dayanıklı tüketim malının marjı ve stok günü aynı ölçütle
  değerlendirilemez; varsayılanlar genel bir ölçüt değildir. Kategorileri ayrı ayrı çalıştırıp eşikleri ayarlayın.
- **Net satış ve SMM aynı esasta olmalı.** İkisi de KDV hariç olmalıdır; iade ve indirimler düşülmüş net satış kullanın.
  Arka marj, ciro primi ve katılım bedelleri SMM'ye yansıtılmadıysa marj olduğundan düşük görünür.
- **Stok tek noktadır.** Dönem sonu stok kullanıldığında stok gün ve devir hızı yaklaşık değerdir; mümkünse ortalama
  stok girin. Sezon öncesi yüklenen stok "yavaş dönen" görünebilir.
- **Pazar payı sizin verinize bağlıdır.** Pazar cirosunu güvenilir bir kaynaktan, aynı dönem ve aynı kapsamla girin.
- **Listeden çıkarma kararı değildir.** Düşük performanslı ürün; sepet tamamlayıcılığı, müşteri sadakati veya
  tedarikçi anlaşması nedeniyle rafta kalabilir.
- **Örnek veri:** Kategoriler, markalar, ürünler ve tutarlar kurgusaldır.

## Testler

- Marj, stok gün, devir hızı, GMROI ve büyüme; stok veya önceki dönem verisi olmayan ürün.
- İşaret sınırları (tam %15, tam 90 gün, tam %20 düşüş), performans sınıfları, satışı olmayan ürün.
- Kategori ve marka toplamları, paylar, pazar payı, ABC sınıfları.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
