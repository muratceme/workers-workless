# Sürücü Çalışma ve Dinlenme Süresi Kontrolü · Kod Bloğu

> Filo Yönetimi › Filo Yöneticisi · Gerisi Bende

Takograf verisinden **günlük ve haftalık sürüş ile dinlenme sürelerini AETR kurallarına göre** kontrol eder, ihlalleri
listeler. İnternete bağlanmaz. Girdi günlük özet satırlarıdır; ham takograf dosyası (.ddd) okumaz.

## Nasıl hesaplar?

| Kontrol | Varsayılan sınır | Parametre |
|---|---|---|
| Günlük sürüş | 9 saat; haftada en çok 2 gün 10 saate kadar uzatılabilir; 10 saatin üzeri her durumda ihlal | `--gunluk-surus`, `--uzatilmis-surus`, `--uzatma-hakki` |
| Kesintisiz sürüş | En çok 4,5 saat | `--kesintisiz-surus` |
| Günlük dinlenme | En az 11 saat; haftada en çok 3 kez 9 saate kadar azaltılabilir; 9 saatin altı ihlal | `--gunluk-dinlenme`, `--azaltilmis-dinlenme`, `--azaltma-hakki` |
| Haftalık sürüş | Pazartesi-pazar toplamı en çok 56 saat | `--haftalik-surus` |
| İki haftalık sürüş | Art arda iki haftanın toplamı en çok 90 saat | `--iki-haftalik-surus` |

- 10 saati aşan gün hem "Günlük sürüş" ihlalidir hem de o haftanın uzatma hakkından düşer.
- Süreler `9,5` (ondalık saat) veya `9:30` (saat:dakika) yazılabilir.

**Örnek:** bir haftada 9:30, 10:30, 9:30, 9:30 ve 8:00 sürüş → ikinci gün 10 saati aştığı için ihlal; üçüncü ve
dördüncü günler haftanın 3. ve 4. uzatması olduğu için ihlal. Önceki hafta 52 saat, bu hafta 47 saat → iki haftalık
99 saat > 90 → ihlal.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --kayitlar takograf_ozeti.xlsx
python main.py --kayitlar takograf_ozeti.xlsx --gunluk-surus 9 --uzatilmis-surus 10 --haftalik-surus 56
```

Sütunlar: Sürücü, Tarih, Sürüş Süresi (zorunlu); En Uzun Kesintisiz Sürüş, Günlük Dinlenme, Plaka. Her sürücü için
gün başına tek satır.

## Çıktı

`cikti/surus_dinlenme_kontrolu.xlsx`:

- `İhlaller`: sürücü, gün veya dönem, kural, gerçekleşen süre, sınır, fark; boş "Alınan Önlem" sütunu.
- `Sürücü Özeti`: kayıtlı gün, toplam sürüş, ihlal sayısı ve ihlal edilen kurallar.
- `Haftalık Toplamlar`: haftalık ve iki haftalık sürüş, uzatılmış gün ve azaltılmış dinlenme sayıları.
- `Günlük Kayıtlar`, `Uyarılar`.

## Dikkat

- **Kaynak teyidi yapılmamıştır.** Varsayılan sınırlar AETR'nin genel sürüş ve dinlenme kurallarına göre seçilmiştir;
  güncel AETR metnini, Karayolları Trafik mevzuatını ve yurt içi taşımalar için geçerli kuralları siz doğrulayın.
- **Kapsam dışı kurallar:** haftalık dinlenme (45 saat, azaltılmış 24 saat ve telafisi), 6 günlük çalışma sonrası
  haftalık dinlenme, molanın süresi ve bölünmesi (45 dk = 15 + 30), bölünmüş günlük dinlenme (3 + 9 saat), ekipli
  sürüş, feribot ve tren istisnaları, sürüş dışı çalışma süreleri kontrol edilmez.
- **Günlük özet yaklaşımıdır.** Gerçek kural 24 saatlik dönemler ve iki dinlenme arasındaki süre üzerinden işler;
  takvim günü özeti gece yarısını aşan seferlerde farklı sonuç verebilir. Kesin değerlendirme için takograf analiz
  yazılımınızın raporu esastır.
- **Azaltılmış dinlenme sayımı haftalıktır.** Kural iki haftalık dinlenme arasındaki dönemi esas alır; kod
  pazartesi-pazar haftasını kullanır.
- **Eksik gün ihlal sayılmaz.** Kayıt olmayan günler dinlenme kabul edilmez; yalnız verdiğiniz günler değerlendirilir.
- **Kişisel veri:** Sürücü adı yerine sicil veya kart numarası kullanabilirsiniz.
- **Örnek veri:** Sürücüler, plakalar ve süreler kurgusaldır.

## Testler

- Süre biçimlerinin okunması.
- Günlük sürüş, uzatma hakkı, kesintisiz sürüş, günlük dinlenme ve azaltma hakkı sınırları; haftanın yenilenmesi.
- Haftalık 56 ve iki haftalık 90 saat sınırları; arada boş hafta.
- Okuma hataları, özel sınırlar ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| AETR sürüş ve dinlenme süreleri (günlük 9 saat, haftalık 56, iki haftalık 90; 4,5 saatte 45 dk mola; günlük 11 saat dinlenme) | [ÇSGB](https://csgb.gov.tr/cgm/tr/haberler/27-07-2021/), [UND](https://www.und.org.tr/files/2024/06/surucu-faaliyetleri-ve-takografa-kaydedilmesi-hakkinda-bilgilendirme.pdf) | Teyit edildi (bakanlık ve sektör kaynağı) |

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
