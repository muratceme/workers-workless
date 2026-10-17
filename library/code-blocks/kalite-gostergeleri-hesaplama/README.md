# Kalite Göstergeleri Hesaplama · Kod Bloğu

> Kalite Yönetimi › Kalite Birim Sorumlusu · Gerisi Bende

Kayıtlardan **SKS kalite göstergelerini** (düşme, bası yarası, enfeksiyon oranı vb.) **formüllerine göre dönemsel**
hesaplar; hedefle ve önceki dönemle karşılaştırır. İnternete bağlanmaz. Hasta bilgisi gerekmez; yalnız sayılar girilir.

## Nasıl hesaplar?

Formüller koda gömülü değildir. Her göstergeyi **gösterge tanım dosyasına** siz yazarsınız:

| Gösterge | Pay | Payda | Çarpan | Hedef | Yön | Ölçü Birimi |
|---|---|---|---|---|---|---|
| Düşme oranı (ÖRNEK TANIM) | Düşen hasta sayısı | Hasta günü | 1000 | 1,00 | Düşük iyi | binde (hasta günü) |

| Hesap | Kural |
|---|---|
| Gösterge değeri | Pay verisi ÷ payda verisi × çarpan |
| Kurum toplamı | Birimlerin pay toplamı ÷ payda toplamı × çarpan (birim oranlarının ortalaması değildir) |
| Hedef durumu | "Düşük iyi": değer ≤ hedef ise Hedefte. "Yüksek iyi": değer ≥ hedef ise Hedefte. Aksi halde Hedef dışı |
| Eğilim | Kurum değeri önceki dönemle karşılaştırılır; yöne göre İyileşti / Kötüleşti / Değişmedi |

- Veri dosyasında aynı dönem, birim ve veri adı birden çok satırda geçerse **toplanır** (haftalık girişler için).
- Bir birimde göstergenin payı veya paydasından yalnız biri girilmişse o birim kurum toplamına katılmaz ve uyarılır.
  **Olay yoksa boş bırakmayın, 0 yazın.**
- Bir birimde göstergenin ne payı ne paydası varsa birim o göstergeyi izlemiyor sayılır; listelenmez.

**Örnek:** iki birimde düşen hasta 3 ve 0, hasta günü 2.000 ve 500 → kurum değeri 3 ÷ 2.500 × 1000 = binde 1,20;
hedef 1,00 ve yön "Düşük iyi" → Hedef dışı. Önceki dönem 0,84 → Kötüleşti.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (son dönem: 2026-09)
python main.py --veriler gosterge_verileri.xlsx --tanimlar gosterge_tanimlari.xlsx --donem 2026-09
```

| Dosya | Sütunlar |
|---|---|
| Veriler | Dönem, Veri, Değer (zorunlu); Birim. Dönem: `2026-09`, `09.2026` veya `Eylül 2026` |
| Gösterge tanımları | Gösterge, Pay, Payda (zorunlu); Çarpan (boşsa 100), Hedef, Yön (Düşük iyi / Yüksek iyi), Ölçü Birimi |

Veri adları iki dosyada aynı yazılmalıdır (büyük-küçük harf ve Türkçe karakter farkı sorun olmaz).

## Çıktı

`cikti/kalite_gostergeleri.xlsx`:

- `Gösterge Özeti`: rapor döneminde kurum değeri, hedef durumu, önceki dönem, eğilim; boş "İyileştirme Faaliyeti" sütunu.
- `Birim Bazında`: rapor döneminde her birimin değeri ve hedef durumu.
- `Dönem Eğilimi`: göstergelerin tüm dönemlerdeki kurum değerleri.
- `Uyarılar`: eksik veriler, hiç girilmemiş veri adları, okunamayan satırlar.

## Dikkat

- **Örnek tanımlar doğrulanmamıştır.** Örnek dosyadaki üç göstergenin formülü, çarpanı ve hedefi yalnız dosya
  biçimini göstermek içindir. Geçerli tanımları Sağlıkta Kalite Standartları (SKS) gösterge kartlarından ve
  kurumunuzun gösterge yönetimi prosedüründen alın; pay ve paydaya hangi hastaların dahil edileceği (dahil etme ve
  dışlama ölçütleri) orada yazar.
- **Veri tanımı hesabın kendisinden önemlidir.** "Hasta günü", "yatan hasta", "risk altındaki hasta" gibi paydalar
  kurumlar arasında farklı sayılabilir; kod yalnız girdiğiniz sayıları böler.
- **Küçük paydada oran oynaktır.** Az yataklı bir birimde tek olay oranı çok yükseltir; birim bazındaki "Hedef dışı"
  sonuçlarını pay ve paydaya bakarak yorumlayın.
- **Eğilim tek dönem farkıdır.** İstatistiksel anlamlılık testi yapılmaz; kalıcı değişim için `Dönem Eğilimi`
  sayfasındaki birkaç dönemi birlikte değerlendirin.
- **Kişisel sağlık verisi:** Hasta adı, T.C. kimlik no veya protokol numarası gerekmez; yalnız sayım sonuçlarını girin.
- **Örnek veri:** Birimler, sayılar, hedefler ve gösterge tanımları kurgusaldır.

## Testler

- Değer, kurum toplamı (oran ortalaması değil), hedef sınırı, "Yüksek iyi" göstergeler, payda 0.
- Eğilim (iyileşme, kötüleşme, değişmeme), rapor dönemi seçimi.
- Eksik veri: yalnız payı veya paydası olan birim, göstergeyi izlemeyen birim.
- Tanım ve veri okuma hataları, aynı verinin satırlarının toplanması, örnek veri uçtan uca.

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
