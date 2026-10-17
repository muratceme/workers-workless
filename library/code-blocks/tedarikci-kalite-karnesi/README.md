# Tedarikçi Kalite Karnesi · Kod Bloğu

> Kalite Güvence ve Kalite Kontrol › Kalite Müdürü · Gerisi Bende

Giriş kalite kontrol kayıtlarından **tedarikçi bazında ret oranını (ppm)** hesaplar, hatalı ve reddedilen parti
sayılarıyla birlikte A / B / C karne notu verir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Ret oranı (ppm) | Σ ret miktarı ÷ Σ gelen miktar × 1.000.000 |
| Ret oranı % | Σ ret miktarı ÷ Σ gelen miktar × 100 |
| Hatalı parti oranı % | Ret miktarı sıfırdan büyük ya da kararı Ret / Şartlı kabul olan parti ÷ toplam parti × 100 |
| Reddedilen parti | Kararı "Ret" olan parti; karar yazılmamışsa tamamı reddedilen parti |
| Karne notu | A: ppm ≤ `--a` (1.000) · B: ppm ≤ `--b` (5.000) · C: üstü |

Parti sayısı `--min` (3) altında olan tedarikçiye not verilmez, "Az veri" yazılır.

**Örnek:** Dört partide 20.000 adet geldi, 50 adet reddedildi → 50 ÷ 20.000 × 1.000.000 = **2.500 ppm**, not **B**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek veri
python main.py --kayitlar giris_kontrol.xlsx
python main.py --kayitlar giris_kontrol.xlsx --a 500 --b 2500 --min 5
```

Zorunlu sütunlar: Tedarikçi, Gelen Miktar. Her satır bir giriş partisidir.

## Çıktı

`cikti/tedarikci_kalite_karnesi.xlsx`:

- `Karne`: ppm'i en yüksek tedarikçiden başlayarak; not, miktarlar, hatalı ve reddedilen parti; son satır toplam.
- `Tedarikçi × Malzeme`: aynı göstergeler malzeme kırılımıyla.
- `Aylar`: aylık ppm seyri.
- `Partiler`, `Uyarılar`.

## Dikkat

- **Not eşikleri politikanızdır.** 1.000 ve 5.000 ppm varsayılanları yalnız başlangıç değeridir; otomotiv gibi
  sektörlerde beklenen düzey çok daha düşük, bazı hammaddelerde daha yüksektir. Eşikleri sözleşmelerinize göre verin.
- **ppm, kontrol yönteminize bağlıdır.** Örnekleme ile kontrol ediyorsanız ret miktarı yalnız örnekte bulunan
  hatayı mı, partinin tamamını mı gösteriyor; bütün tedarikçilerde aynı tanımı kullanın. Hatta ve müşteride ortaya
  çıkan tedarikçi kaynaklı hatalar bu veride yoktur.
- **Miktar ağırlıklı ortalamadır.** Çok adetli ucuz parça (cıvata) ile az adetli kritik parçayı (döküm gövde) aynı
  ppm ölçeğinde kıyaslamak yanıltır; malzeme kırılımına bakın.
- **Az partide not oynaktır.** Tek kötü parti notu belirler; `--min` ile eşiği dönem uzunluğunuza göre ayarlayın.
- **Yalnız kalite boyutudur.** Zamanında teslimat, fiyat, düzeltici faaliyetlere cevap süresi ve denetim sonuçları
  bu karnede yoktur; tedarikçi değerlendirmesinin bütünü değildir.
- **Örnek veri:** Tedarikçiler, malzemeler ve miktarlar kurgusaldır.

## Testler

- ppm, ret oranı, hatalı ve reddedilen parti; not eşiklerinin sınırı (tam eşikteki ppm üst nota girer).
- En az parti koşulu, karar yazılmamış partinin ret miktarına göre değerlendirilmesi.
- Tedarikçi × malzeme ve ay kırılımı, ret miktarı gelenden büyük satır, okuma ve komut satırı.

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
