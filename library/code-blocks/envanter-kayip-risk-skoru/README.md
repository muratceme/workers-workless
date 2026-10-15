# Envanter Kayıp Risk Skoru · Kod Bloğu

> Kayıp Önleme › Kayıp Önleme Uzmanı · Workers / Workless

Sayım farkı, fire, iade ve iptal göstergelerinden **mağaza bazında kayıp risk skoru** üretir ve denetim önceliği
çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Oran % | Gösterge tutarı ÷ net satış × 100 |
| Gösterge puanı | Oran ÷ eşik × 50, en çok 100 (oran eşikteyse 50; eşiğin iki katı ve üzerindeyse 100) |
| Risk skoru | Σ(ağırlık × puan) ÷ Σ(ağırlık) |
| Risk sınıfı | Skor ≥ 70 Yüksek; ≥ 40 Orta; altı Düşük |
| Denetim önceliği | Skora göre büyükten küçüğe sıra |

| Gösterge | Varsayılan eşik (satışın %'si) | Varsayılan ağırlık |
|---|---|---|
| Sayım eksiği | 1 | 40 |
| Fire | 1 | 20 |
| İade | 5 | 20 |
| İptal | 2 | 20 |

- Sayım farkında yalnız **eksik** (negatif fark) kayıp sayılır; fazla çıkan mağazada kayıp 0 alınır ve bilgi notu düşülür.
- Verisi olmayan gösterge o mağaza için hesaptan çıkar; ağırlıklar kalan göstergelere yeniden dağıtılır ve uyarı verilir.
- "Baskın gösterge", skora en çok katkı yapan göstergedir.

**Örnek:** satış 1.000.000; sayım eksiği 25.000 (%2,5 → puan 100), fire 8.000 (%0,8 → 40), iade 60.000 (%6 → 60),
iptal 30.000 (%3 → 75) → skor 0,4 × 100 + 0,2 × 40 + 0,2 × 60 + 0,2 × 75 = 75 → Yüksek.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --magazalar magaza_gostergeleri.xlsx --sayim-esik 1,5 --iade-esik 8 --agirlik 50,10,20,20
```

Sütunlar: Mağaza, Net Satış (zorunlu); Bölge, Sayım Farkı (eksik negatif yazılır), Fire, İade, İptal. Hepsi aynı
döneme ait tutar olmalıdır.

## Çıktı

`cikti/kayip_risk_skoru.xlsx`:

- `Denetim Önceliği`: öncelik sırası, skor, sınıf, baskın gösterge, eşiği aşan göstergeler, oranlar ve puanlar.
- `Göstergeler`: eşik, ağırlık, zincir ortalaması, eşiği aşan mağaza sayısı.
- `Uyarılar`: eksik göstergeler, sayım fazlası, okunamayan değerler.

## Dikkat

- **Eşik ve ağırlıklar varsayılandır, standart değildir.** Kayıp oranı sektöre, ürün grubuna ve mağaza formatına göre
  çok değişir (taze gıdada fire, hazır giyimde iade doğal olarak yüksektir). Önce `Göstergeler` sayfasındaki zincir
  ortalamalarına bakıp eşikleri kendi verinize göre belirleyin.
- **Skor suçlama değildir.** Yüksek skor; hırsızlık kadar sayım hatası, mal kabul hatası, yanlış fiyat veya süreç
  eksikliğinden de kaynaklanabilir. Sonuç yalnız denetim sırasını belirler; kişiler hakkında çıkarım yapmayın.
- **Tutarlar aynı esasta olmalı.** Sayım farkı ve fire maliyet bedeliyle, satış satış bedeliyle tutuluyorsa oranlar
  düşük görünür; hepsini aynı bedelle girin veya eşikleri buna göre ayarlayın.
- **Dönem aynı olmalı.** Sayım farkı iki sayım arasındaki dönemi kapsar; satış ve diğer göstergeleri aynı dönem için alın.
- **İade ve iptal her zaman kayıp değildir.** Bunlar dolaylı risk göstergeleridir; personel bazında iptal, iade
  saatleri gibi ayrıntılar bu pakette yoktur.
- **Örnek veri:** Mağazalar ve tutarlar kurgusaldır.

## Testler

- Oran, puan, skor ve sınıf; tam eşikteki mağaza; eksik göstergede ağırlıkların yeniden dağıtılması.
- Özel eşik ve ağırlık, zincir ortalaması, öncelik sırası.
- Okuma hataları, sayım fazlası ve örnek veri uçtan uca.

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
