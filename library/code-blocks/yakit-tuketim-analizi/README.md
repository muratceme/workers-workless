# Yakıt Tüketim Analizi · Kod Bloğu

> Filo Yönetimi › Filo Yöneticisi · Workers / Workless

Yakıt kartı ve kilometre verisinden **araç ve sürücü bazında 100 km tüketimini** hesaplar, sapmaları işaretler.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Dilim | Aynı aracın art arda iki alımı arası: mesafe = km sayacı farkı, yakıt = ikinci alımın litresi |
| Araç tüketimi (L/100 km) | İlk alım hariç litre toplamı ÷ (son km − ilk km) × 100 |
| Km başına maliyet | İlk alım hariç tutar toplamı ÷ toplam mesafe |
| Hedefe göre sapma % | (Tüketim − hedef) ÷ hedef × 100; `--tolerans` (%10) üzerindeyse "Hedefin üzerinde" |
| Şüpheli dilim | Dilim tüketimi araç ortalamasının `--dilim-tolerans` (%30) üzerinde |
| Sürücü tüketimi | Sürücünün yakıt aldığı dilimlerin litre toplamı ÷ mesafe toplamı × 100 |

İlk alım yalnız başlangıç noktasıdır: o yakıtın kaç kilometrede harcandığı bilinmediği için hesaba girmez.
Ek kontroller: km sayacı geri gitmiş veya değişmemiş alım; depo kapasitesinden fazla litre.

**Örnek:** km 100.000 → 103.000 arasında 300 + 310 + 480 = 1.090 L → 1.090 ÷ 3.000 × 100 = 36,33 L/100 km; hedef 30 →
sapma %21,11 → Hedefin üzerinde. Son dilim 48,0 L/100 km, ortalamanın %30'undan fazla üzerinde → şüpheli.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --alimlar yakit_alimlari.xlsx --araclar araclar.xlsx --tolerans 8 --dilim-tolerans 25
```

| Dosya | Sütunlar |
|---|---|
| Yakıt alımları | Plaka, Tarih, Km (sayaç), Litre (zorunlu); Tutar, Sürücü |
| Araçlar (isteğe bağlı) | Plaka (zorunlu); Hedef Tüketim (L/100 km), Depo Kapasitesi (L), Araç Tipi |

## Çıktı

`cikti/yakit_tuketim.xlsx`:

- `Araç Bazında`: mesafe, yakıt, tüketim, hedef, sapma, km başına maliyet, şüpheli dilim sayısı; filo toplamı.
- `Sürücü Bazında`: sürücüye yazılan dilimlerin toplamı ve tüketimi.
- `Dilimler`: şüpheli olanlar başta olmak üzere her dilim.
- `Uyarılar`: km sayacı sorunları, depo kapasitesini aşan alımlar, okunamayan satırlar.

## Dikkat

- **Depo her alımda dolduruluyor varsayılır.** Kısmi alımlarda dilim tüketimi bir dilimde düşük, sonrakinde yüksek
  çıkar ve "şüpheli" işaretlenebilir. Araç toplamı uzun dönemde yine de anlamlıdır; kısa dönemde yanıltır.
- **Sapma suistimal kanıtı değildir.** Yük, güzergah, eğim, trafik, mevsim, rölanti, lastik ve arıza tüketimi
  değiştirir. Sonuç inceleme önceliği verir; kişiler hakkında çıkarım yapmadan önce telematik ve sefer verisine bakın.
- **Sürücü ataması yaklaşıktır.** Dilim, alımı yapan sürücüye yazılır; dilim içinde araç başka sürücüye geçtiyse
  bilinemez.
- **Km verisi kritik.** Elle girilen km sayacı hatalıysa bütün hesap bozulur; sayacı geri giden alımlar atlanır.
- **Hedef sizin verinizdir.** Örnek dosyadaki hedefler "örnek hedef" olarak işaretlidir; üretici değeri veya kendi
  geçmiş ortalamanızı kullanın. AdBlue ve soğutucu ünite yakıtı ayrı tutulmalıdır.
- **Örnek veri:** Plakalar, sürücüler ve tutarlar kurgusaldır.

## Testler

- Araç tüketimi, km başına maliyet, hedef sapması sınırı (tam %10), şüpheli dilim, sürücü toplamları.
- Km sayacı geri gitmiş alım, depo kapasitesini aşan alım, tek alımı olan araç, plaka yazım farkı.
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
