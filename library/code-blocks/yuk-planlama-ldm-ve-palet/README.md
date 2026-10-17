# Yük Planlama (LDM ve Palet) · Kod Bloğu

> Karayolu Operasyon › Karayolu Operasyon Uzmanı · Gerisi Bende

Yük ölçüleri, palet sayısı ve istiflenebilirliğe göre **yükleme metresini (LDM), ağırlığı ve araç tipini** hesaplar.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Zemin yeri | Adet ÷ istif katı, yukarı yuvarlanır (istif katı 1: istiflenemez) |
| LDM | Zemin yeri × boy (m) × en (m) ÷ araç genişliği (`--genislik`, varsayılan 2,40 m) |
| Euro palet eşdeğeri | LDM ÷ 0,4 (120 × 80 cm palet = 0,4 LDM) |
| Hacim, ağırlık | Adet × boy × en × yükseklik; adet × birim ağırlık |
| Araç seçimi | Toplam LDM ≤ yükleme uzunluğu, toplam ağırlık ≤ kapasite ve her kalem sığan **en kısa** araç |
| Sığmazsa | En büyük araçtan gereken sayı: en büyük(LDM ÷ uzunluk, ağırlık ÷ kapasite), yukarı yuvarlanır |

Kalem kontrolü: dar kenar ≤ araç iç genişliği, uzun kenar ≤ yükleme uzunluğu, istifli yükseklik ≤ iç yükseklik.

**Örnek:** 10 euro palet (istiflenemez) → 10 × 1,2 × 0,8 ÷ 2,4 = 4,00 LDM. 8 palet 2 kat istifli → 4 zemin yeri → 1,60 LDM.
2 sandık 300 × 200 cm → 2 × 6 ÷ 2,4 = 5,00 LDM. Toplam 10,60 LDM.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --yukler yukler.xlsx --araclar arac_tipleri.xlsx --genislik 2,45
```

| Dosya | Sütunlar |
|---|---|
| Yükler | Adet, Boy (cm), En (cm) (zorunlu); Yük, Yükseklik (cm), Ağırlık (kg, birim başına), İstif Katı (sayı veya Evet = 2) |
| Araç tipleri | Araç Tipi, Yükleme Uzunluğu (m), Kapasite (kg) (zorunlu); İç Genişlik (m), İç Yükseklik (m) |

## Çıktı

`cikti/yuk_plani.xlsx`:

- `Yük Planı`: kalem bazında zemin yeri, LDM, ağırlık, hacim; toplamlar ve önerilen araç.
- `Araç Uygunluğu`: her araç tipi için uygunluk, doluluk oranları ve uygun değilse nedeni.
- `Uyarılar`: okunamayan satırlar, eksik ölçüler, tek araca sığmayan yük.

## Dikkat

- **LDM alan yaklaşımıdır.** Kalemlerin yan yana gerçekten sığıp sığmadığı, yükleme sırası, ağırlık dağılımı, dingil
  yükü ve yük emniyeti (bağlama) hesaplanmaz; yükleme planını ayrıca yapın.
- **Araç ölçüleri sizin verinizdir.** Örnek dosyadaki ölçü ve kapasiteler "örnek ölçü" olarak işaretlidir; kendi
  filonuzun ruhsat ve üstyapı değerlerini yazın. Yasal ağırlık ve boyut sınırlarını kod bilmez.
- **İstif katı beyana dayanır.** İstiflenebilirlik, alt kalemin taşıyabileceği yük ve ambalaj dayanımı kontrol edilmez.
- **Birden çok araca bölmez.** Yük tek araca sığmıyorsa yalnız yaklaşık araç sayısını verir.
- **Ölçü birimleri:** boy, en, yükseklik santimetre; ağırlık kilogram; araç ölçüleri metre.
- **Örnek veri:** Yükler ve araç tipleri kurgusaldır.

## Testler

- LDM: istifsiz, istifli (tek sayıda adet), büyük sandık, 33 palet, farklı genişlik.
- Araç seçimi sınırları (LDM, ağırlık, yükseklik, genişlik), tek araca sığmayan yük.
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
