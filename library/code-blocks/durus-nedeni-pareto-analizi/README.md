# Duruş Nedeni Pareto Analizi · Kod Bloğu

> Üretim › Üretim Mühendisi · Gerisi Bende

Makine duruş kayıtlarını **neden ve süreye göre Pareto analizine** tabi tutar: toplam duruşun büyük kısmını yaratan
az sayıdaki nedeni ve makineyi öne çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Süre (dk) | `Süre (dk)` sütunu; boşsa bitiş − başlangıç saati (bitiş başlangıçtan küçükse gece yarısını geçtiği kabul edilir) |
| Pay % | Nedenin toplam süresi ÷ toplam duruş süresi × 100 (`--esas adet` ise duruş adedi üzerinden) |
| Kümülatif % | Nedenler büyükten küçüğe sıralandığında payların birikimli toplamı |
| A sınıfı | Kendisinden önceki kümülatif pay `--esik` (80) yüzdesinin altında olan neden |
| Ortalama duruş (dk) | Toplam süre ÷ duruş adedi |
| Kapsam | Plansız duruşlar; `--planli-dahil` verilirse planlı duruşlar da girer |

`Tür` sütunu "Planlı" ile başlayan kayıtlar planlı, diğerleri (boş dahil) plansız sayılır.

**Örnek:** Toplam 635 dakikalık plansız duruşun 300 dakikası mekanik arıza → pay **%47,24**. Malzeme bekleme (100 dk)
ve kalıp değişimi (95 dk) eklendiğinde kümülatif **%77,95**; dördüncü neden %80 sınırını geçiren kalemdir ve o da
A sınıfına girer.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --duruslar duruslar.xlsx
python main.py --duruslar duruslar.xlsx --esas adet --esik 70 --planli-dahil --uzun 60
```

Zorunlu sütunlar: Makine, Duruş Nedeni; ayrıca `Süre (dk)` ya da `Başlangıç` ve `Bitiş`.

## Çıktı

`cikti/durus_pareto_analizi.xlsx`:

- `Neden Pareto`: sıra, sınıf, adet, süre, pay, kümülatif pay, ortalama duruş.
- `Makine Pareto`: aynı tablo makine bazında; her makinenin en büyük duruş nedeni.
- `Özet`: toplamlar, A sınıfı neden sayısı, analize girmeyen planlı duruşlar, vardiya kırılımı.
- `Duruşlar`, `Uyarılar`.

## Dikkat

- **Pareto kayıt kalitesi kadar iyidir.** Neden kodları tutarsız yazılmışsa ("arıza", "mekanik arıza", "makine
  arızası") aynı neden üç satıra bölünür ve sıralama değişir. Kod yalnız büyük-küçük harf ve Türkçe karakter
  farkını yok sayar; nedenleri sizin için birleştirmez.
- **"Diğer" büyükse analiz yanıltır.** Ayrıştırılmamış duruşların payı %10'u aşarsa uyarı üretilir.
- **Süre ile adet farklı öykü anlatır.** Süreye göre sıralama uzun arızaları, adede göre sıralama sık tekrarlayan
  kısa duruşları öne çıkarır; ikisine de bakın (`--esas`).
- **Kod kök nedeni bulmaz.** "Mekanik arıza" bir belirtidir; neden-sonuç analizi (5 neden, balık kılçığı) ve
  aksiyon belirleme sizin işinizdir.
- **Çok kısa duruşlar genellikle kayda girmez.** Mikro duruşlar ve hız kayıpları bu veride yoktur; toplam kaybı
  görmek için ekipman etkinliği (OEE) hesabına bakın.
- **Gece yarısı varsayımı:** bitiş saati başlangıçtan küçükse duruşun ertesi güne geçtiği kabul edilir; 24 saati
  aşan duruşlar için süreyi dakika olarak yazın.
- **Örnek veri:** Makineler, nedenler ve süreler kurgusaldır.

## Testler

- Pay, kümülatif pay ve A sınıfı sınırı; süre ve adet esasına göre sıralama.
- Saatten süre hesabı (gece yarısını geçen duruş), planlı duruşların kapsam dışı kalması ve `--planli-dahil`.
- Makine bazında en büyük neden, vardiya kırılımı, "Diğer" ve nedeni yazılmamış duruş uyarıları, okuma ve komut satırı.

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
