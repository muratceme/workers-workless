# Zaman Etüdü ve Standart Süre · Kod Bloğu

> Endüstri Mühendisliği ve Süreç Geliştirme › Endüstri Mühendisi · Gerisi Bende

Kronometre ölçümlerinden **performans ve tolerans payıyla** eleman ve operasyon bazında normal süreyi, standart
süreyi ve saatlik standart çıktıyı hesaplar; gözlem sayısının yeterliliğini kontrol eder. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Normal süre (gözlem)   = gözlenen süre × performans % ÷ 100
Eleman normal süresi   = gözlemlerin normal sürelerinin ortalaması × frekans
Standart süre          = normal süre × (1 + tolerans % ÷ 100)
Operasyon süresi       = elemanların toplamı
Saatlik standart çıktı = 3600 ÷ operasyon standart süresi (sn)
Gerekli gözlem sayısı  = (k ÷ e)² × (n·Σx² − (Σx)²) ÷ (Σx)²        (yukarı yuvarlanır)
```

- **Frekans:** her parçada yapılan eleman için 1; on parçada bir yapılan için `1/10` (veya 0,1).
- **Gerekli gözlem:** varsayılan k = 2 ve e = %5 ile çarpan 1600'dür (yaklaşık %95 güven, ±%5 hata). Gözlem sayısı
  bundan azsa eleman işaretlenir.
- **Örnek:** dört elemanın normal süreleri 5,00 + 13,20 + 3,60 + 2,00 = 23,80 sn; %12 toleransla standart süre
  26,66 sn; saatlik standart çıktı 135,05 adet.
- Performans ve tolerans satırda yoksa `--performans` (varsayılan 100) ve `--tolerans` (varsayılan 0) kullanılır.
- Süreler saniyedir; virgül ve nokta ondalık ayırıcıdır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (tolerans %12)
python main.py --gozlemler gozlemler.xlsx --tolerans 12 --vardiya-dk 450
python main.py --hata 10 --guven-k 2                   # ±%10 hata için gerekli gözlem
```

Sütunlar: Eleman, Süre (sn) (zorunlu); Operasyon, Performans %, Frekans, Tolerans %. Her satır bir kronometre ölçümüdür.

## Çıktı

`cikti/standart_sure.xlsx`:

- `Standart Süreler`: operasyon bazında gözlenen, normal ve standart süre; saatlik ve vardiya çıktısı.
- `Elemanlar`: eleman bazında ortalama, en kısa / en uzun, performans, frekans, tolerans, gerekli gözlem sayısı.
- `Uyarılar`: gözlemi yetersiz elemanlar, okunamayan satırlar.

## Dikkat

- **Performans değerlendirmesi yargıdır.** Kod, etüdü yapanın verdiği tempoyu aynen kullanır; yanlış derecelendirme
  standart süreyi aynı oranda bozar.
- **Tolerans payını siz belirlersiniz.** Kişisel ihtiyaç, yorgunluk ve kaçınılmaz gecikme payları işin niteliğine
  göre değişir; varsayılan 0'dır. Toplu sözleşme veya işyeri uygulamasındaki oranları kullanın.
- **Tolerans normal süreye eklenir.** Payı çalışma gününün yüzdesi olarak tanımlıyorsanız (normal süre ÷ (1 − pay))
  sonuç farklıdır; oranı buna göre dönüştürün.
- **Aykırı gözlemler atılmaz.** Hatalı veya olağandışı ölçümleri dosyadan siz çıkarın ve nedenini not edin.
- **Makine süreleri:** makine kontrollü elemanlarda performans 100 girilmelidir.
- **Ücret ve prim** sistemlerinde kullanmadan önce standart sürelerin ilgili taraflarca doğrulanması gerekir.
- **Örnek veri:** Operasyon ve ölçümler kurgusaldır.

## Testler

- Eleman ve operasyon süreleri, frekanslı eleman, saatlik çıktı.
- Gerekli gözlem sayısı (yetersiz, tek gözlem, değişkenliği sıfır olan eleman, farklı hata payı).
- Varsayılan performans ve tolerans, gözlem bazında performans, eleman bazında tolerans.
- Sayı okuma, okuma hataları ve örnek veri uçtan uca.

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
