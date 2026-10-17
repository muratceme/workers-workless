# Gümrük Evrak Tutarlılık Kontrolü · Kod Bloğu

> Lojistik ve Taşımacılık › Gümrük Operasyon › Gümrük Operasyon Uzmanı · Gerisi Bende

Fatura, çeki listesi, menşe belgesi ve taşıma belgesindeki ürün, miktar, ağırlık, değer ve taraf bilgilerini
karşılaştırır; **beyanname öncesi düzeltilmesi gereken uyumsuzlukları** listeler. İnternete bağlanmaz.

## Ne kontrol eder?

| Kontrol | Önem |
|---|---|
| Alıcı, gönderici veya fatura numarası belgeler arasında farklı | Yüksek |
| Kap adedi farklı | Yüksek |
| Fatura hesabı: miktar × birim fiyat ≠ tutar; kalem toplamı ≠ fatura toplamı | Yüksek |
| Kalem: faturada olup çeki / menşe belgesinde olmayan, veya tersi | Yüksek |
| Kalem: miktar farkı | Yüksek |
| Kalem: GTİP'in ilk 6 hanesi farklı | Yüksek |
| Kalem: menşe ülke farklı | Yüksek |
| Çeki listesinde net ağırlık brütten büyük | Yüksek |
| Brüt / net ağırlık veya diğer sayısal alanlar toleransın üzerinde farklı | Orta |
| Çeki kalem toplamı belge bilgisinden farklı | Orta |
| Tarih alanları farklı | Orta |
| Teslim şekli Incoterms 2020 kurallarından biri değil | Orta |
| FAS / FOB / CFR / CIF yalnız deniz ve iç su yolu taşımacılığı içindir; taşıma belgesi CMR, AWB gibi başka bir türse uyarılır | Orta |
| GTİP 12 hane değil (Türkiye GTİP 12 hanedir; 6–8 haneli kodlar beyannamede tamamlanmalı) | Bilgi |

**Karşılaştırma kuralları:**
- **Sayısal alanlar:** Belgelerin çoğunluğundaki değer referans alınır; toleransın (`--tolerans`, varsayılan
  %0,5) dışında kalan belge işaretlenir.
- **Unvanlar:** "GmbH", "Ltd.", "A.Ş." gibi ekler ve noktalama farkı sayılmaz.
- **Kalemler:** Ürün koduyla, kod yoksa tanımla eşleştirilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                      # örnek ihracat evrakı (bilerek hatalı)
python main.py --klasor ./sevkiyat_evraki
python main.py --klasor ./sevkiyat_evraki --tolerans 0.01
```

Klasördeki dosyalar (`.csv` / `.xlsx`):

| Dosya | İçerik |
|---|---|
| `belge_bilgileri` | İlk sütun alan adı (Gönderici, Alıcı, Fatura No, Fatura Tarihi, Döviz, Fatura Toplamı, Teslim Şekli, Kap Adedi, Brüt Ağırlık, Net Ağırlık, Taşıma Belgesi Türü / No, Varış Ülkesi, Menşe Ülke…); sonraki sütunlar belgeler |
| `fatura…` | Ürün Kodu, Tanım, GTİP, Miktar, Birim, Birim Fiyat, Tutar, Menşe |
| `ceki…` / `packing…` | Ürün Kodu, Tanım, Miktar, Kap Adedi, Net Ağırlık, Brüt Ağırlık |
| `mense…` / `eur1…` / `atr…` | Ürün Kodu, Tanım, GTİP, Miktar, Menşe Ülke |
| `tasima…` / `cmr…` / `konsimento…` | İsteğe bağlı kalem listesi |

Belge türü dosya adından anlaşılır. PDF evraktaki bilgileri bu tablolara aktarmanız gerekir.

## Çıktı

`gumruk_evrak_kontrolu.xlsx`:
- `Bulgular`: önem, konu, belge, açıklama; boş "Düzeltme / Not" sütunu.
- `Belge Bilgileri`: alan × belge; uyumsuz hücreler kırmızı, uyumlu olanlar yeşil.
- `Kalem Karşılaştırma`: ürün bazında miktarlar ve notlar.
- Her kalem belgesinin kendi sayfası.

## Dikkat

- **Kapsam:** Tutarlılık kontrolüdür. GTİP sınıflandırmasının, menşe kurallarının, vergi ve ticaret politikası
  önlemlerinin doğruluğunu denetlemez; bunları güncel mevzuat ve gümrük müşaviriyle teyit edin.
- **Incoterms uyarısı:** Incoterms 2020'nin genel kuralına göredir; taraflar sözleşmede farklı uzlaşmış olabilir.
- **Örnek veri:** Örnek evrak, firmalar ve tutarlar kurgusaldır.

## Testler

Şunlar test edilir:
- Örnek evraktaki 9 uyumsuzluğun tamamı ve belge bazında işaretleme.
- Tolerans; tutarlı evrak ve Incoterms kuralları; Excel renkleri.
- Komut satırı.

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
