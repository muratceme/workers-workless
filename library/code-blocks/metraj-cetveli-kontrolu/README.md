# Metraj Cetveli Kontrolü · Kod Bloğu

> Teknik Ofis › Teknik Ofis Mühendisi · Gerisi Bende

Metraj cetvelindeki **hesap satırlarını (boyut × adet) yeniden hesaplar, poz toplamlarını ve mükerrer satırları**
kontrol eder. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kontrol | Kural |
|---|---|
| Hesaplanan miktar | Benzer × adet × boy × en × yükseklik. Boş bırakılan çarpan 1 sayılır |
| Minha | "Minha" sütunu Evet olan veya miktarı eksi yazılmış satır eksi alınır |
| Hesap hatası | \|Hesaplanan − yazılan miktar\| > `--tolerans` (varsayılan 0,01) |
| Poz toplamı | Pozun satırlarının hesaplanan toplamı; yazılan toplamla ve verilirse icmal miktarıyla karşılaştırılır |
| Mükerrer satır | Aynı poz, aynı tanım ve aynı çarpanlarla birden fazla satır |

Sayılarda **virgül de nokta da ondalık ayırıcıdır** (`3,25` = `3.25`; `12.500` = 12,5). Binlik ayırıcı yalnız
`1.250,50` yazımında tanınır.

**Örnek:** benzer 2, adet 3, boy 4,00, en 0,25, yükseklik 0,50 → 2 × 3 × 4 × 0,25 × 0,5 = 3,00; cetvelde 2,50 yazıyor →
hesap hatası, fark 0,50.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --metraj metraj_cetveli.xlsx --icmal icmal.xlsx --tolerans 0,005
```

| Dosya | Sütunlar |
|---|---|
| Metraj cetveli | Poz No (zorunlu); Tanım, Benzer, Adet, Boy, En, Yükseklik, Miktar, Minha |
| İcmal (isteğe bağlı) | Poz No, Miktar |

## Çıktı

`cikti/metraj_kontrolu.xlsx`:

- `Satır Kontrolü`: hatalı satırlar başta; yazılan ve hesaplanan miktar, fark, mükerrer işareti; boş "Düzeltme" sütunu.
- `Poz Toplamları`: yazılan ve hesaplanan toplam, hatalı ve mükerrer satır sayısı, icmal karşılaştırması.
- `Uyarılar`: okunamayan çarpanlar, çarpanı olmayan satırlar, icmalde olmayan pozlar.

## Dikkat

- **Aritmetiği denetler, metrajı değil.** Ölçülerin projeyle uyumunu, doğru pozun seçildiğini, boşlukların düşülme
  kuralını, fire ve bindirme paylarını ve ölçü birimini değerlendirmez.
- **Mükerrer işareti bir sorudur.** Aynı ölçülerde iki ayrı imalat olabilir; tanımı ayırt edici yazın (aks, kat, mahal).
- **Formüllü satırlar desteklenmez.** Üçgen, daire, yamuk gibi özel alan ve hacim formülleri çarpım olarak
  yazılamıyorsa sonucu tek çarpan olarak girin; satır "çarpanı olmayan" ise yalnız toplamda kullanılır.
- **Yazılmamış miktar hata sayılmaz.** Hesaplanan değer poz toplamına girer; satır "Miktar yazılmamış" olarak listelenir.
- **Örnek veri:** Pozlar ve ölçüler kurgusaldır.

## Testler

- Sayı okuma (virgül ve nokta), çarpım, boş çarpan, minha, tolerans sınırı, yazılmamış miktar.
- Poz toplamları, mükerrer satır, icmal karşılaştırması, icmalde veya cetvelde olmayan poz.
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
