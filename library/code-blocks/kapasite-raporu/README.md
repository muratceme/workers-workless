# Kapasite Raporu · Kod Bloğu

> Bilgi Teknolojileri › Sistem Uzmanı · Gerisi Bende

Disk, bellek ve işlemci kullanım ölçümlerinizden **büyüme hızını** çıkarır ve "bu disk ne zaman dolar?" sorusuna
tarihle yanıt verir. Eşiğe yaklaşan kaynakları öne alır; satın alma ve temizlik planı için tek tablo üretir.
İnternete bağlanmaz, sunuculara erişmez.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Günlük büyüme | Ölçümlere en küçük kareler yöntemiyle uydurulan doğrunun eğimi (birim ÷ gün) |
| Kullanım oranı | Son ölçüm ÷ son kapasite |
| Eşiğe kalan gün | (Kapasite × eşik % − son ölçüm) ÷ günlük büyüme, yukarı yuvarlanır |
| Dolmaya kalan gün | (Kapasite − son ölçüm) ÷ günlük büyüme, yukarı yuvarlanır |
| R² | Doğrunun ölçümleri açıklama oranı; 1 tam doğrusal, 0,5 altı "eğilim zayıf" |

Durumlar: **Eşik aşıldı** (oran ≥ eşik), **Kritik** (eşiğe 30 gün veya daha az), **Uyarı** (90 gün veya daha az),
**Normal**, **Büyüme yok** (eğim sıfır veya eksi), **Yetersiz veri** (farklı tarihlerde 3'ten az ölçüm).

**Örnek:** 2.000 GB disk, haftada 35 GB artışla 1.675 GB'a gelmiş → günlük büyüme 5 GB. %90 eşiği 1.800 GB;
(1.800 − 1.675) ÷ 5 = **25 gün** sonra eşiğe, (2.000 − 1.675) ÷ 5 = **65 gün** sonra tam doluluğa ulaşır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri, durum tarihi 12.10.2026
python main.py --olcumler olcumler.csv
python main.py --olcumler olcumler.csv --esik 85 --kritik 45 --uyari 120 --min-olcum 5
```

Zorunlu sütunlar: Tarih, Sistem, Kullanılan, Kapasite. Diğerleri: Kaynak, Birim. Her satır bir kaynağın bir
tarihteki ölçümüdür. Yüzdeyle izlenen kaynaklarda (bellek, işlemci) Kullanılan sütununa yüzdeyi, Kapasite
sütununa 100 yazın.

| Seçenek | Anlamı | Varsayılan |
|---|---|---|
| `--esik` | Kullanım eşiği, % | 90 |
| `--kritik`, `--uyari` | Eşiğe kalan gün sınırları | 30, 90 |
| `--min-olcum` | Eğilim için gereken en az ölçüm | 3 |
| `--bayat` | Son ölçüm bu kadar günden eskiyse uyarılır | 14 |

## Çıktı

`cikti/kapasite_raporu.xlsx`:

- `Kapasite Özeti`: kaynak bazında durum, kullanım oranı, günlük ve 30 günlük büyüme, R², eşiğe ve dolmaya kalan
  gün ile tahmini tarihler. En acil kaynak en üsttedir.
- `Uyarılar`: eşiği aşan ve yaklaşan kaynaklar, zayıf eğilim, eski veya yetersiz ölçüm, değişen kapasite.

## Dikkat

- **Tahmin doğrusal bir uzatmadır.** Geçmişteki hızın aynen süreceği varsayılır; toplu veri yükleme, arşivleme,
  temizlik, yeni proje veya mevsimsellik gibi kırılmaları kod bilemez. R² düşükse tarihe güvenmeyin.
- **Bellek ve işlemci için sınırlıdır.** Bu kaynaklar anlık dalgalanır; ortalama veya tepe değerlerle düzenli
  ölçüm girmezseniz eğilim anlamsız çıkar. Yöntem en çok disk ve depolama alanına uygundur.
- **Ölçümleri siz sağlarsınız.** Kod sistemlere bağlanıp ölçüm almaz; izleme aracınızdan dışa aktardığınız
  veriyi okur. Aynı güne birden çok ölçüm varsa sonuncusu kullanılır.
- **Eşikler tercihtir.** %90, 30 ve 90 gün genel başlangıç değerleridir; tedarik sürenize göre ayarlayın.
- **Örnek veri:** Sistemler ve ölçümler kurgusaldır.

## Testler

- Eğim ve R² (elle hesaplanmış), eşiğe ve dolmaya kalan gün, yukarı yuvarlama, tarih hesabı.
- Durum sınırları (30 ve 31, 90 ve 91 gün), eşik aşımı, büyüme yok, yetersiz veri, aynı güne çift ölçüm.
- Eski ölçüm, değişen kapasite, okuma hataları ve komut satırı.

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
