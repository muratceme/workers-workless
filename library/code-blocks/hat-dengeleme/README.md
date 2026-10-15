# Hat Dengeleme · Kod Bloğu

> Endüstri Mühendisliği ve Süreç Geliştirme › Endüstri Mühendisi · Workers / Workless

Operasyon süreleri ve öncelik ilişkilerinden, hedef çevrim süresine göre **istasyon atamasını, hat verimini ve denge
gecikmesini** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Çevrim süresi (C)      = --cevrim;  verilmezse (--sure-dk × 60) ÷ --talep
Teorik asgari istasyon = ⌈ toplam operasyon süresi ÷ C ⌉
Hat verimi %           = toplam operasyon süresi ÷ (istasyon sayısı × C) × 100
Denge gecikmesi %      = 100 − hat verimi
Düzgünlük endeksi      = √ Σ (en dolu istasyonun süresi − istasyon süresi)²
```

**Atama kuralı (konum ağırlığı, `rpw`):**

1. Her operasyonun konum ağırlığı = kendi süresi + kendisinden sonra gelmek zorunda olan tüm operasyonların süresi.
2. Operasyonlar konum ağırlığına göre büyükten küçüğe sıralanır.
3. Açık istasyona, öncüllerinin tümü atanmış ve kalan süreye sığan ilk operasyon eklenir.
4. Sığan operasyon kalmayınca yeni istasyon açılır.

`--yontem lct` ile sıralama operasyon süresine göre yapılır (en uzun süre önce). İki yöntemi de deneyip istasyon
sayısı az, düzgünlük endeksi düşük olanı seçebilirsiniz.

**Örnek:** 10 operasyon, toplam 274 sn, çevrim 80 sn → teorik asgari 4 istasyon. Bulunan atama: {A, B, E}, {C, F},
{D, G, H}, {I, J}; hat verimi 274 ÷ 320 = %85,63.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (çevrim 80 sn)
python main.py --operasyonlar operasyonlar.xlsx --cevrim 80
python main.py --operasyonlar operasyonlar.xlsx --talep 330 --sure-dk 440 --yontem lct
```

Sütunlar: Operasyon (kod), Süre (zorunlu); Açıklama, Öncül (virgülle ayrılmış operasyon kodları; yoksa boş).
Süreler ile çevrim süresi aynı birimde olmalıdır; `--talep` ve `--sure-dk` kullanıyorsanız süreler saniye olmalıdır.

## Çıktı

`cikti/hat_dengeleme.xlsx`:

- `Özet`: çevrim süresi, teorik ve bulunan istasyon sayısı, hat verimi, denge gecikmesi, düzgünlük endeksi, saatlik çıktı.
- `İstasyonlar`: istasyon bazında operasyonlar, süre, boş süre ve doluluk.
- `Operasyonlar`: süre, öncüller, konum ağırlığı ve atandığı istasyon.
- `Uyarılar`: okunamayan veya tekrar eden operasyonlar.

Süresi çevrim süresinden uzun operasyon, listede olmayan öncül veya öncelik döngüsü varsa hesap yapılmaz, hata mesajı verilir.

## Dikkat

- **Sezgisel yöntemdir.** Bulunan çözüm en az istasyonlu çözüm olmayabilir; "bulunan" istasyon sayısı teorik
  asgariye eşitse daha iyisi yoktur.
- **Yalnız süre ve öncelik kısıtı.** İstasyonun fiziksel yeri, ortak ekipman, operatör becerisi, aynı istasyonda
  bulunmaması gereken işler, çift taraflı hat ve paralel istasyon dikkate alınmaz.
- **Standart süreler doğru olmalı.** Tolerans payı eklenmiş standart süreleri kullanın; ölçüm hatası atamayı değiştirir.
- **Tek model.** Karışık modelli hatlarda ağırlıklı ortalama süreyle çalıştırmak yaklaşık sonuç verir.
- **Örnek veri:** Operasyonlar ve süreler kurgusaldır.

## Testler

- Konum ağırlıkları, istasyon ataması, verim, denge gecikmesi ve düzgünlük endeksinin elle hesaplanmış değerleri.
- Farklı çevrim süreleri ve iki yöntemde öncelik ihlali ve çevrim aşımı olmadığı.
- Çevrimden uzun operasyon, eksik öncül, döngü; okuma hataları ve örnek veri uçtan uca.

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
