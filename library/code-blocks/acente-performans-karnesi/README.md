# Acente Performans Karnesi · Kod Bloğu

> Satış ve Acente Yönetimi › Bölge Satış Yöneticisi · Gerisi Bende

Acente bazında **üretim, büyüme, hasar/prim, tahsilat ve yenileme** göstergelerini hesaplar, puana çevirir ve
acenteleri karşılaştırmalı bir karnede sıralar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül | Ağırlık | 0 puan | 100 puan |
|---|---|---|---|---|
| Hedef gerçekleşme % | Üretim ÷ hedef | 25 | %50 ve altı | %100 ve üstü |
| Büyüme % | (Üretim − önceki dönem üretim) ÷ önceki dönem üretim | 15 | %0 ve altı | %50 ve üstü |
| Hasar/prim % | Gerçekleşen hasar ÷ kazanılmış prim | 25 | %100 ve üstü | %40 ve altı |
| Tahsilat oranı % | Tahsilat ÷ tahakkuk (vadesi gelen prim) | 20 | %70 ve altı | %100 |
| Yenileme oranı % | Yenilenen poliçe ÷ yenilenecek poliçe | 15 | %50 ve altı | %90 ve üstü |

```
Gösterge puanı = (değer − sıfır puan değeri) ÷ (tam puan değeri − sıfır puan değeri) × 100     (0-100 arasına sınırlanır)
Karne puanı    = Σ(gösterge puanı × ağırlık) ÷ Σ(hesaplanabilen göstergelerin ağırlığı)
```

- Arada kalan değerler doğrusal puanlanır: hedef gerçekleşme %75 → 50 puan; hasar/prim %70 → 50 puan.
- **Hesaplanamayan gösterge** (ör. yeni acentede önceki dönem üretimi yok) karneden çıkarılır, ağırlığı diğerlerine
  dağılır ve "Hesaplanamayan" sütununda belirtilir.
- **Not:** puan ≥ 80 A, ≥ 65 B, ≥ 50 C, altı D (`--notlar 80,65,50`).
- **En zayıf gösterge:** puanı 50'nin altında olan en düşük puanlı gösterge.
- Tablodaki ağırlık ve sınırlar örnek varsayılanlardır; `--olcutler` dosyasıyla değiştirilir. Ağırlığı 0 olan
  gösterge karneye girmez.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri, varsayılan ölçütler
python main.py --acenteler acente_verisi.xlsx --olcutler olcutler.xlsx --notlar 85,70,55
```

| Dosya | Sütunlar |
|---|---|
| Acente verisi | Acente, Üretim (zorunlu); Bölge, Önceki Dönem Üretim, Hedef, Kazanılmış Prim, Gerçekleşen Hasar, Tahakkuk, Tahsilat, Yenilenecek Poliçe, Yenilenen Poliçe |
| Ölçütler (isteğe bağlı) | Gösterge, Ağırlık, Sıfır Puan, Tam Puan |

## Çıktı

`cikti/acente_karnesi.xlsx`:

- `Karne`: genel ve bölge içi sıra, göstergeler, karne puanı, not, en zayıf gösterge; boş "Aksiyon" sütunu.
- `Gösterge Puanları`: her göstergenin 0-100 puanı (50'nin altı işaretli).
- `Bölge Özeti`: bölge toplamlarından hesaplanan göstergeler ve ortalama karne puanı.
- `Ölçütler`: kullanılan ağırlık ve sınırlar.
- `Uyarılar`: okunamayan değerler, mükerrer acenteler.

## Dikkat

- **Ölçütler şirket politikasıdır.** Ağırlıklar ve puan sınırları sektör standardı değildir. Yüksek enflasyon
  dönemlerinde nominal büyüme sınırını (varsayılan %50) özellikle gözden geçirin.
- **Küçük acentelerde oranlar oynaktır.** Birkaç poliçelik portföyde tek bir büyük hasar H/P oranını bozar; karneyi
  üretim payıyla birlikte okuyun.
- **Aynı dönem, aynı tanım.** Tüm sütunlar aynı dönemi kapsamalı; hasar tutarı ödenen + muallak ise bunu tüm
  acentelerde aynı şekilde kullanın.
- **Karar aracı değildir.** Komisyon, kategori değişikliği veya sözleşme feshi gibi kararlar yalnız bu puana
  dayandırılmamalıdır; acentenin portföy yapısı ve bölge koşulları ayrıca değerlendirilmelidir.
- **Örnek veri:** Acenteler ve tutarlar kurgusaldır.

## Testler

- Beş gösterge ve karne puanının elle hesaplanmış değerleri.
- Puan sınırları (0 ve 100'e sınırlama, düşük olanın iyi olduğu gösterge).
- Hesaplanamayan gösterge, ağırlığı 0 olan gösterge, not sınırları.
- Sıralama, bölge özeti, okuma hataları ve örnek veri uçtan uca.

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
