# Eksper Performans Analizi · Kod Bloğu

> Hasar Yönetimi › Hasar Müdürü · Workers / Workless

Eksperleri **rapor süresi, tespit tutarı ile ödeme arasındaki fark ve itiraz oranına göre** karşılaştırır; genel
değerlerden belirgin sapanları işaretler, raporu bekleyen dosyaları listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Rapor süresi (gün) | Rapor tarihi − atama tarihi; eksper bazında ortalama ve ortanca |
| Süre aşımı % | Süresi `--sla` (15) günü aşan dosya ÷ raporlanan dosya × 100 |
| Bekleyen dosya | Rapor tarihi boş; bekleme = durum tarihi − atama tarihi |
| Tespit–ödeme farkı % | (Σ ödenen − Σ tespit) ÷ Σ tespit × 100 |
| Mutlak sapma % | Σ \|ödenen − tespit\| ÷ Σ tespit × 100 |
| İtiraz oranı % | İtiraz edilen dosya ÷ raporlanan dosya × 100 |

Tutar göstergelerine yalnız raporlanmış, tespit ve ödeme tutarı birlikte yazılmış dosyalar girer. Raporlanan dosyası
`--min` (3) ve üstünde olan eksperde, bir gösterge genel değerin (1 + `--sapma` %20) katını aşıyorsa eksper
işaretlenir: "Süre uzun", "Tespit–ödeme farkı yüksek", "İtiraz yüksek" (itiraz için ayrıca en az 2 itiraz aranır).
Daha az dosyası olan eksper "Az veri" olarak gösterilir.

**Örnek:** Bir eksperin dört raporu 22, 20, 26 ve 16 günde gelmiş → ortalama **21 gün**; genel ortalama 13,17 gün,
eşik 15,8 gün → "Süre uzun". Tespit toplamı 278.000, ödenen 239.000 → fark **−%14,03**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                  # örnek veri, durum tarihi 12.10.2026
python main.py --dosyalar ekspertiz_dosyalari.xlsx
python main.py --dosyalar ekspertiz_dosyalari.xlsx --sla 10 --min 20 --sapma 30 --tarih 31.10.2026
```

Zorunlu sütunlar: Eksper, Atama Tarihi. Ödeme henüz yapılmadıysa Ödenen Tutar hücresini boş bırakın; sıfır yazarsanız
ödemesiz kapanmış sayılır ve farka girer.

## Çıktı

`cikti/eksper_performans_analizi.xlsx`:

- `Eksperler`: bütün göstergeler ve değerlendirme; son satır genel değerler.
- `Bekleyen Raporlar`: raporu gelmemiş dosyalar, bekleme süresine göre.
- `Dosyalar`: dosya bazında süre, fark ve itiraz.
- `Uyarılar`.

## Dikkat

- **Fark, eksperin hatası anlamına gelmez.** Ödeme tespitten; muafiyet, eksik sigorta, sovtaj, kusur oranı,
  sigortalıyla uzlaşma ya da şirket kararı nedeniyle de farklı olabilir. Kod nedeni bilmez; işaretlenen eksperin
  dosyalarını tek tek inceleyin.
- **Dosyalar eşit zorlukta değildir.** Büyük yangın dosyası ile küçük kasko dosyasının süresi karşılaştırılamaz.
  Farklı branş ya da tutar dilimleri için dosyayı ayırıp ayrı çalıştırmak daha adil sonuç verir.
- **Az dosyada oranlar oynaktır.** `--min` ile değerlendirme eşiğini dosya hacminize göre yükseltin.
- **Genel değer, eksperin kendisini de içerir.** Dosyaların çoğu tek eksperdeyse o eksper genel değeri belirler ve
  sapma görünmez.
- **Süre takvim günüdür.** Eksperin elinde olmayan beklemeler (sigortalıdan evrak, servis onarımı) süreden
  düşülmez.
- **Kapsam dışı:** rapor kalitesi puanlaması, eksper atama dengesi, ek rapor ve revizyon sayıları, mevzuattaki
  ekspertiz süre ve ücret kuralları.
- **Örnek veri:** Eksperler, dosyalar ve tutarlar kurgusaldır.

## Testler

- Ortalama ve ortanca süre, süre aşımı, bekleyen dosya, fark ve mutlak sapma, itiraz oranı, ücret.
- Değerlendirme eşikleri, en az dosya koşulu, ödemesi boş dosyanın tutar göstergelerine girmemesi.
- Rapor tarihi atamadan önce, geciken rapor uyarısı, okuma ve komut satırı.

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
