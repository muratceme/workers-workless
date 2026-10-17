# Stok ABC-XYZ Analizi · Kod Bloğu

> Üretim › Üretim Planlama · Gerisi Bende

Stokları iki eksende sınıflandırır ve her sınıf için **stok politikası** önerir. İnternete bağlanmaz.

| Eksen | Ölçüt | Varsayılan sınırlar |
|---|---|---|
| **ABC** (değer) | Tüketim değeri = miktar × birim maliyet, Pareto sıralaması | A ≤ %80, B ≤ %95, C kalan (kümülatif pay) |
| **XYZ** (talep değişkenliği) | Değişim katsayısı CV = standart sapma ÷ ortalama | X ≤ 0,5 · Y ≤ 1,0 · Z > 1,0 |

- **Sınırı aşan kalem:** Kümülatif payı sınırı aşan ilk kalem üst sınıfta kalır; örneğin %78'den %83'e çıkaran
  kalem A olur.
- **Yeni kalemler:** İlk tüketimden önceki dönemler sıfır talep sayılmaz. Bu sayede yeni kalem yapay olarak
  "düzensiz" görünmez.
- **Hareketsiz:** Hiç tüketimi olmayan kalemler ayrı listelenir.
- **Yetersiz geçmiş:** `--en-az-donem` (varsayılan 6) dönemden az geçmişi olan kalemler XYZ yerine "?"
  işaretiyle gösterilir.

Dokuz sınıfın (AX … CZ) politika önerileri rapordadır. Örnek: AX → sürekli yenileme ve düşük emniyet stoğu;
AZ → siparişe göre alım; CZ → stok tutma gerekliliğini sorgulama.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                    # örnek 12 aylık tüketimle
python main.py --girdi tuketim.xlsx --maliyet maliyetler.xlsx
python main.py --girdi tuketim.xlsx --abc 70 90 --xyz 0.25 0.5     # kendi sınırlarınız
```

ERP'den alınan **stok çıkış hareketleri** (Tarih, Stok Kodu, Miktar) doğrudan kullanılabilir; aylara toplanır.

## Çıktı

`Matris` (3×3 tablo, politika listesi, Pareto grafiği) · `Kalemler` · `Dönem Tüketimi` · `Bilgi`

## Testler

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
