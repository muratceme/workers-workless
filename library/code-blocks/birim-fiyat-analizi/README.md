# Birim Fiyat Analizi · Kod Bloğu

> Teknik Ofis › Teknik Ofis Mühendisi · Workers / Workless

Malzeme, işçilik ve makine rayiçlerinden **poz bazında birim fiyat analizi** yapar ve **kâr-genel gider dahil fiyatı**
oluşturur. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Girdi tutarı | Miktar × birim fiyat |
| Analiz toplamı | Pozun malzeme + işçilik + makine + diğer girdi tutarları |
| Kâr ve genel gider | Analiz toplamı × `--kar-orani` % (varsayılan 25) |
| Birim fiyat | Analiz toplamı + kâr ve genel gider |

- **Birim fiyat kaynağı:** analiz satırında yazılıysa o kullanılır; boşsa rayiç dosyasından rayiç numarasıyla alınır.
  İkisi de yoksa girdi 0 fiyatla alınır ve yüksek önemli uyarı verilir.
- **Tür** sütunu Malzeme, İşçilik veya Makine olarak tanınır; diğerleri (nakliye vb.) "Diğer" sayılır.
- Miktarlarda virgül de nokta da ondalık ayırıcıdır (`0,25` = `0.25`). Fiyatlarda nokta binlik ayırıcıdır (`4.000,00`).

**Örnek:** malzeme 1.726 + işçilik 1.350 + makine 400 = 3.476; %25 kâr ve genel gider 869 → birim fiyat 4.345.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --analiz analiz_girdileri.xlsx --rayicler rayicler.xlsx --kar-orani 25
```

| Dosya | Sütunlar |
|---|---|
| Analiz girdileri | Poz No, Miktar (zorunlu); Poz Tanımı, Ölçü Birimi, Tür, Rayiç No, Girdi, Birim, Birim Fiyat |
| Rayiçler (isteğe bağlı) | Rayiç No, Fiyat (zorunlu); Tanım, Birim |

## Çıktı

`cikti/birim_fiyat_analizi.xlsx`:

- `Poz Özeti`: malzeme, işçilik, makine ve diğer toplamları, analiz toplamı, kâr ve genel gider, birim fiyat, paylar.
- `Analiz Detayı`: her girdinin miktarı, fiyatı, tutarı, poz içindeki payı ve fiyat kaynağı.
- `Uyarılar`: bulunamayan rayiçler, birim uyuşmazlıkları, rayiçten farklı fiyatlar, aynı girdinin iki kez yazılması.

## Dikkat

- **Resmi birim fiyat değildir.** Kod hiçbir kurumun birim fiyat kitabını, analiz numarasını, analiz miktarını veya
  güncel rayicini bilmez. Örnek dosyalardaki miktar ve fiyatlar kurgusaldır; kendi analiz ve rayiç listenizi kullanın.
- **Oran sözleşmeye bağlıdır.** %25 kâr ve genel gider varsayılandır; işinizin şartnamesindeki oranı `--kar-orani`
  ile verin. Nakliye, KDV ve fiyat farkı ayrı hesaplanır.
- **Birim uyumu sizin sorumluluğunuzda.** Analiz miktarı rayicin birimine göre yazılmalıdır (kg ile ton, saat ile
  gün karışırsa sonuç yanlış olur). Birimler farklı yazıldığında uyarı verilir ama çevrilmez.
- **Eksik fiyatlı pozun birim fiyatı düşük çıkar.** `Poz Özeti`ndeki "Fiyatı Eksik Girdi" sütununa bakın.
- **Örnek veri:** Pozlar, analiz miktarları ve rayiçler kurgusaldır.

## Testler

- Poz toplamı, tür payları, kâr ve genel gider oranı, kuruş yuvarlama.
- Birden çok poz, fiyatı eksik girdi, aynı girdinin iki kez yazılması, tür ve miktar okuma.
- Rayiç eşleştirme, okuma hataları ve örnek veri uçtan uca.

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
