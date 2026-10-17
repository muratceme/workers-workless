# Kontrol Testi Örneklem Seçimi · Kod Bloğu

> İç Kontrol ve Risk Yönetimi › İç Kontrol Uzmanı · Gerisi Bende

İşlem popülasyonundan kontrol testi için **tekrarlanabilir örneklem** seçer: aynı dosya ve aynı tohumla her
çalıştırmada aynı kalemler çıkar, böylece seçimi çalışma kâğıdında belgeleyebilir ve başkası yeniden üretebilir.
Eşit olasılıklı, tutar ağırlıklı ve risk tabakalı üç yöntem sunar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Yöntem | Kural |
|---|---|
| `rastgele` | Her kaleme tohumlu rastgele bir sayı verilir; en küçük sayıya sahip N kalem seçilir (eşit olasılık) |
| `tutar` | Aralık = toplam tutar ÷ örneklem sayısı. Tutarı aralığa eşit ya da büyük kalemler **kesin** seçilir; kalanlar için aralık yeniden hesaplanır ve rastgele bir başlangıçtan itibaren her aralıkta bir kalem seçilir |
| `risk` | Her risk tabakasına (Yüksek / Orta / Düşük) kalem sayısı × ağırlık oranında pay verilir; tabaka içinde rastgele seçilir |
| Kapsama % | Örneklem ÷ popülasyon × 100 (adet ve tutar bazında) |

**Örnek (tutar yöntemi):** 30 kalem, toplam 1.708.600, örneklem 10 → aralık 170.860. 530.000 ve 310.000 tutarlı kalemler
kesin seçilir. Kalan 28 kalem için aralık 868.600 ÷ 8 = 108.575 → 150.000 ve 125.000 de kesin seçilir. Kalan 26 kalem
için aralık 593.600 ÷ 6 = 98.933,33; bu aralıkla 6 kalem daha seçilir.

**Örnek (risk yöntemi):** 5 yüksek, 10 orta, 15 düşük riskli kalem; ağırlıklar 3;2;1 → puanlar 15, 20, 15 → 10'luk
örneklem **3 + 4 + 3** olarak dağılır. Yüksek riskli kalemlerin %60'ı, düşük risklilerin %20'si test edilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri: rastgele, 10 kalem, tohum 2026
python main.py --populasyon populasyon.xlsx --adet 25 --tohum 2026
python main.py --populasyon populasyon.xlsx --yontem tutar --adet 40 --tohum 7
python main.py --populasyon populasyon.xlsx --yontem risk --adet 30 --agirlik "4;2;1"
```

Zorunlu sütun: İşlem No. Tutar yöntemi için Tutar, risk yöntemi için Risk sütunu gerekir.

## Çıktı

`cikti/kontrol_testi_orneklem.xlsx`:

- `Örneklem`: seçilen kalemler popülasyondaki sırasıyla; seçim nedeni ve sizin dolduracağınız test sonucu, bulgu ve
  testi yapan sütunları.
- `Özet`: popülasyon ve örneklem adedi ve tutarı, kapsama; seçim parametreleri (kaynak dosya, yöntem, tohum, aralık,
  başlangıç) ve risk tabakası dağılımı.
- `Uyarılar`: mükerrer işlem numarası, tutarı olmayan kalemler, örneklemin popülasyonun tamamı olması.

## Dikkat

- **Örneklem büyüklüğünü siz belirlersiniz.** Kontrolün sıklığına, riskine ve kabul edilebilir sapmaya göre kaç
  kalem test edileceği metodolojinizin konusudur; kod bir büyüklük önermez.
- **Kod sonuç çıkarmaz.** Testte bulunan hataların popülasyona yansıtılması, güven düzeyi ve hata oranı tahmini bu
  paketin kapsamında değildir. Risk tabakalı ve tutar ağırlıklı seçimlerde olasılıklar eşit olmadığı için basit
  oran yansıtması yanlış olur.
- **Tekrarlanabilirlik dosyaya bağlıdır.** Aynı örneklem için dosyanın içeriği ve **satır sırası** da aynı olmalıdır;
  popülasyonu seçimden önce sabitleyin ve dosyayı saklayın. Tohumu çalışma kâğıdına yazın; tohumu sonuca bakıp
  değiştirmek seçimi yanlı yapar.
- **Popülasyonun tamlığını siz doğrularsınız.** Kod dosyada olmayan işlemi seçemez; popülasyonu kaynak sistemle
  (adet ve tutar) mutabık kılın.
- **Tutar yönteminde eksi ve sıfır tutarlar seçilemez.** İade ve ters kayıtları ayrı bir popülasyon olarak
  değerlendirin.
- **Örnek veri:** İşlemler ve tutarlar kurgusaldır.

## Testler

- Tekrarlanabilirlik (aynı tohum aynı örneklem, farklı tohum farklı örneklem), örneklem sayısı, popülasyondan büyük istek.
- Tutar yöntemi: kesin seçimler ve aralığın yeniden hesaplanması, sistematik seçim (elle verilen başlangıçla), eksi tutar.
- Risk yöntemi: tabakalara dağıtım (tam bölünen, kalanlı, kapasitesi dolan tabaka), belirsiz risk; okuma ve komut satırı.

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
