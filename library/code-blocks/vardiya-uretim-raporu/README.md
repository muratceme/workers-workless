# Vardiya Üretim Raporu · Kod Bloğu

> Üretim › Vardiya Amiri · Gerisi Bende

Vardiya bazında **üretim adedi, hedef, duruş ve fire** kayıtlarından vardiya raporu üretir: vardiya ve hat özeti,
duruş nedenleri ve kopyalanmaya hazır rapor metni. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Sağlam adet         = üretilen adet − fire adet
Hedef gerçekleşme % = sağlam adet ÷ hedef adet × 100
Fire %              = fire adet ÷ üretilen adet × 100
Kullanılabilirlik % = (planlı süre − planlı duruş − plansız duruş) ÷ (planlı süre − planlı duruş) × 100
```

- **Üretilen adet toplamdır** (sağlam + fire). Fire üretilenden büyükse uyarı verilir.
- **Kullanılabilirlik** yalnız `Planlı Süre (dk)` girildiyse hesaplanır. Planlı duruşlar (mola, planlı bakım, kalıp
  değişimi) paydadan düşülür; plansız duruşlar kayıp sayılır.
- **İşaretler:** gerçekleşme < `--hedef-esik` (varsayılan %90) → "Hedefin altında"; fire % > `--fire-esik`
  (varsayılan %3) → "Fire yüksek".
- Duruşun türü `Planlı` ile başlamıyorsa plansız sayılır.
- **Örnek:** hedef 4.000, üretilen 3.900, fire 78 → sağlam 3.822, gerçekleşme %95,55, fire %2,00. Planlı süre 480 dk,
  planlı duruş 30 dk, plansız duruş 15 dk → kullanılabilirlik 435 ÷ 450 = %96,67.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --uretim uretim.xlsx --duruslar duruslar.xlsx --tarih 05.10.2026 --hedef-esik 95 --fire-esik 2
```

| Dosya | Sütunlar |
|---|---|
| Üretim | Tarih, Vardiya, Üretilen Adet (zorunlu); Hat, Ürün, Hedef Adet, Fire Adet, Planlı Süre (dk) |
| Duruşlar (isteğe bağlı) | Tarih, Vardiya, Süre (dk) (zorunlu); Hat, Duruş Nedeni, Tür (`Planlı` / `Plansız`) |

Vardiya ve hat adları iki dosyada aynı yazılmalıdır. `--tarih` verilmezse dosyadaki tüm günler raporlanır.

## Çıktı

`cikti/vardiya_uretim_raporu.xlsx`:

- `Vardiya Özeti`: vardiya ve gün toplamı; gerçekleşme, fire, duruş, kullanılabilirlik ve işaretler.
- `Hat Bazında`: aynı göstergeler hat ve ürün bazında.
- `Duruş Nedenleri`: plansız ve planlı duruşların süreye göre sıralı listesi, pay ve birikimli yüzde.
- `Vardiya Duruşları`: vardiya bazında duruş nedenleri.
- `Uyarılar`: okunamayan satırlar, üretim kaydı olmayan vardiya veya hatlara yazılmış duruşlar.

`cikti/vardiya_raporu.txt`: her vardiya için kısa rapor metni ve boş "Vardiya amiri notu" satırı.

## Dikkat

- **OEE değildir.** Performans bileşeni için ideal çevrim süresi gerekir; bu paket yalnız adet ve süre toplar.
  Kullanılabilirlik, OEE'nin tek bir bileşenidir.
- **Hedef sizin planınızdır.** Kod hedefin gerçekçi olup olmadığını değerlendirmez; ürün değişimi olan vardiyalarda
  hedefi ürün bazında ayrı satırlarla girin.
- **Farklı ürünlerin adetleri toplanır.** Vardiya toplamı, birim süreleri farklı ürünlerde yanıltıcı olabilir; hat
  ve ürün satırlarına bakın.
- **Eşikler işletmeye özeldir.** %90 ve %3 yalnız varsayılandır; sektör standardı değildir.
- **Kayıt kalitesi:** Yazılmayan duruş raporda görünmez; duruş süreleri toplamının vardiya süresini aşmadığını kontrol edin.
- **Örnek veri:** Hatlar, ürünler ve adetler kurgusaldır.

## Testler

- Hat ve vardiya göstergelerinin elle hesaplanmış değerleri; planlı / plansız duruş ayrımı.
- İşaret sınırları (tam %90 ve tam %3 işaretlenmez), özel eşikler, hedefi veya süresi olmayan satırlar.
- Gün toplamı, duruş nedenleri sıralaması, eşleşmeyen duruş uyarıları; rapor metni.
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
