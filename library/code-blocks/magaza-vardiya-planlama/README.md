# Mağaza Vardiya Planlama · Kod Bloğu

> Mağaza Operasyonları › Mağaza Müdür Yardımcısı · Gerisi Bende

Saatlik müşteri trafiği tahmini, personel izinleri ve çalışma süresi kurallarına göre **haftalık vardiya planı
taslağı** çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

1. **Saatlik ihtiyaç** = en büyük(`--min-personel`, beklenen müşteri ÷ `--musteri-basina`, yukarı yuvarlanır).
2. **Vardiya net süre** = (bitiş − başlangıç) − mola. Ara dinlenmesi çalışma süresinden sayılmaz.
3. **Atama (sezgisel):** günler en yoğundan başlanarak ele alınır. Her adımda karşılanmamış saatleri en çok kapatan
   vardiya seçilir (eşitlikte kısa olan) ve kurallara uyan, haftalık kalan saati en çok olan personele verilir
   (eşitlikte izin günü çok olana). İhtiyaç bitene veya uygun personel kalmayana kadar sürer.
4. **Kurallar** (hepsi parametreyle değişir):

   | Kural | Varsayılan |
   |---|---|
   | İzin gününde atama yapılmaz; günde tek vardiya | — |
   | Haftalık net çalışma ≤ personelin sınırı | `--haftalik-saat` 45 |
   | Vardiyanın net süresi ≤ günlük sınır | `--gunluk-saat` 11 |
   | İki vardiya arası dinlenme ≥ | `--dinlenme-saat` 11 |
   | Haftada en çok çalışılan gün | `--max-gun` 6 |

5. **Mola kontrolü:** vardiyadaki mola; vardiya süresi 4 saate kadar 15 dk, 4 saatten 7,5 saate kadar 30 dk, 7,5
   saatten fazla 60 dk'nın altındaysa uyarı verilir.

**Örnek:** saatte 45 müşteri, personel başına 20 müşteri → 45 ÷ 20 = 2,25 → 3 personel. Sabah 10:00–18:00, mola 60 dk →
net 7 saat; haftada 6 vardiya 42 saat eder ve 45 saat sınırının içindedir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --trafik saatlik_trafik.xlsx --personel personel.xlsx --vardiyalar vardiyalar.xlsx --musteri-basina 15 --min-personel 3
```

| Dosya | Sütunlar |
|---|---|
| Trafik | Gün (Pazartesi … Pazar veya tarih), Saat (10 veya 10:00), Müşteri. Yalnız açık saatleri yazın |
| Personel | Personel (zorunlu); Haftalık Saat (boşsa sınır değeri), İzin Günleri (virgülle: `Salı, Pazar`) |
| Vardiyalar | Vardiya, Başlangıç, Bitiş (zorunlu); Mola (dk) |

## Çıktı

`cikti/vardiya_plani.xlsx`:

- `Vardiya Planı`: personel × gün tablosu (vardiya, İzin veya boş); toplam net saat, sınır, kalan saat, çalışılan gün.
- `Gün Özeti`: gün başına vardiya sayısı; ihtiyaç, atanan, eksik ve fazla adam-saat; eksik kalan saatler.
- `Saatlik Kapsama`: her saat için beklenen müşteri, ihtiyaç, atanan ve fark.
- `Uyarılar`: karşılanamayan ihtiyaç, asgari sürenin altındaki molalar, tanınmayan izin günleri.

## Dikkat

- **Plan taslaktır; yasal uygunluğu siz doğrularsınız.** Varsayılan sınırlar İş Kanunu'ndaki genel çalışma süresi
  kurallarına göre seçilmiştir, ancak bu pakette kaynak teyidi yapılmamıştır. Güncel mevzuatı, toplu iş sözleşmesini
  ve iş sözleşmelerini kontrol edin.
- **Kapsam dışı kurallar:** fazla çalışma ve ücreti, resmi tatil ve genel tatil çalışması, gece çalışması sınırları,
  denkleştirme, kısmi süreli sözleşme ayrıntıları, genç, gebe ve emziren çalışanlara ilişkin kurallar hesaba katılmaz.
- **Hafta tek başına değerlendirilir.** Önceki haftanın son vardiyası ile bu haftanın ilk vardiyası arasındaki
  dinlenme ve peş peşe çalışılan gün sayısı kontrol edilmez.
- **Sezgiseldir.** En az vardiyayla en iyi kapsamayı garanti etmez; uygun bir çözüm varken eksik bırakabilir. Eksik
  çıkarsa vardiya kalıbı eklemeyi veya izinleri kaydırmayı deneyin.
- **Mola saati planlanmaz.** Personel vardiya boyunca mağazada sayılır; molada olduğu saat kapsamadan düşülmez.
  Vardiya kalıpları üst üste bindiği saatlerde fazla personel görünmesi olağandır.
- **Personel başına müşteri mağazaya özeldir.** 20 bir standart değildir; kasa, reyon ve depo görevlerine göre kendi
  değerinizi belirleyin. Görev ve yetkinlik ayrımı (kasiyer, satış danışmanı) yapılmaz.
- **Gece yarısını aşan vardiyalar desteklenmez.**
- **Kişisel veri:** Personel adı yerine sicil veya kod kullanabilirsiniz.
- **Örnek veri:** Trafik, personel ve vardiyalar kurgusaldır.

## Testler

- Gün ve saat çözümleme, asgari mola basamakları.
- İhtiyaç hesabı, temel atama, izin, kısa vardiya tercihi, günde tek vardiya.
- Haftalık saat sınırı, en çok çalışma günü, vardiyalar arası dinlenme (önceki ve sonraki gün), günlük sınırı aşan vardiya.
- Örnek veride tüm saatlerin karşılanması ve kuralların sağlanması; okuma hataları.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |
| 2429 sayılı Ulusal Bayram ve Genel Tatiller Hakkında Kanun | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.2429.pdf) | Resmî metinle karşılaştırılan maddeler: 1, 2 |

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
