# Şube Hedef-Gerçekleşme Panosu · Kod Bloğu

> Şube Bankacılığı › Şube Müdürü · Workers / Workless

Mevduat, kredi, kart, sigorta gibi kalemlerin **aylık hedeflerini portföy yöneticisi bazında** gerçekleşmelerle
karşılaştırır ve tek panoda toplar. Ayın neresinde olunduğunu iş günü üzerinden hesaba katar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Gerçekleşme % | Gerçekleşen ÷ hedef |
| Olması gereken | Hedef × geçen iş günü ÷ ayın toplam iş günü |
| Sapma | Gerçekleşen − olması gereken |
| Ay sonu tahmini | Gerçekleşen ÷ geçen iş günü × toplam iş günü (aynı tempo sürerse) |
| Kalan günlük gereken | (Hedef − gerçekleşen) ÷ kalan iş günü |
| Karne puanı | Σ ağırlık × min(gerçekleşme %, tavan) ÷ Σ ağırlık |

| Durum | Koşul |
|---|---|
| Hedef tamam | Gerçekleşen ≥ hedef |
| Yolunda | Ay sonu tahmini ≥ hedef |
| Risk | Ay sonu tahmini ≥ hedefin %90'ı |
| Geride | Diğer |

- **İş günü:** Pazartesi–Cuma. Resmî tatiller `--tatiller` dosyasından düşülür. Rapor tarihi dahil sayılır.
- **Karne puanı:** Kalem başına gerçekleşme `--tavan` (varsayılan %120) ile sınırlanır; eksi gerçekleşme 0 sayılır.
  Hiç ağırlık girilmemişse kalemler eşit ağırlıklıdır.
- **Şube toplamı:** Kalem bazında kişilerin hedef ve gerçekleşmeleri toplanır.
- Gerçekleşme satırları rapor tarihine kadar toplanır. Net azalış (ör. mevduat çıkışı) eksi tutarla girilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --hedefler hedefler.xlsx --gerceklesmeler gerceklesmeler.xlsx --tarih 16.10.2026 --tatiller tatiller.csv
```

| Dosya | Sütunlar |
|---|---|
| Hedefler | Portföy Yöneticisi, Kalem, Hedef, Birim, Ağırlık, Ay (isteğe bağlı; `2026-10`) |
| Gerçekleşmeler | Tarih, Portföy Yöneticisi, Kalem, Gerçekleşen |
| Tatiller (isteğe bağlı) | Her satırda bir tarih |

"Ay" sütunu doluysa yalnız rapor tarihinin ayına ait hedefler alınır.

## Çıktı

`cikti/hedef_gerceklesme_panosu.xlsx`:

- `Pano`: kişi × kalem gerçekleşme yüzdeleri, karne puanı ve sıralama. Renkler ayın geçen bölümüne göredir.
- `Detay`: her kişi ve kalem için tüm göstergeler ve durum.
- `Şube Toplamı`: kalem bazında şube sonucu, en iyi ve en zayıf kişi.
- `Uyarılar`: hedefi olmayan gerçekleşme, okunamayan satır, ağırlık sorunları.

## Dikkat

- **Ay sonu tahmini doğrusal bir uzatmadır.** Ay sonu yoğunlaşması, büyük tekil işlemler veya vade dönüşleri hesaba
  katılmaz.
- **Bakiye kalemleri:** Mevduat ve kredi hedefi bakiye artışıysa gerçekleşme satırlarına net değişimi girin. Kod bakiye
  farkı hesaplamaz.
- **Karne puanı prim hesabı değildir.** Prim ve performans sistemleri bankanızın kurallarına göre işler; tavan, taban
  ve ağırlıkları kendi sisteminize göre ayarlayın.
- **Yarım gün tatiller** tam iş günü sayılır.
- **Örnek veri:** Kişiler, hedefler ve gerçekleşmeler kurgusaldır.

## Testler

- İş günü sayımı (Ekim 2026: 21 iş günü, 16 Ekim'e kadar 12).
- Formüller, durum eşikleri, karne puanı (tavan, eksi gerçekleşme, ağırlıksız kalemler), şube toplamı.
- Okuma hataları, dönem filtresi; örnek veri uçtan uca.

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
