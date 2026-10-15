# Hastane Operasyon Panosu · Kod Bloğu

> Hastane Yönetimi › Hastane Müdürü · Workers / Workless

Bölümlerin günlük **poliklinik, acil, yatış, ameliyat, dolu yatak ve gelir** verilerini tek panoda toplar; doluluk,
poliklinikten yatış oranı, hasta başına gelir ve hedef gerçekleşmesini bölüm ve gün bazında gösterir. İnternete ve
HBYS'ye bağlanmaz; dışa aktarılmış günlük tabloyu okur.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Yatak doluluk % | Σ dolu yatak (gün sonu sayımı) ÷ Σ yatak sayısı × 100; ikisi de yazılmış günler üzerinden |
| Poliklinikten yatış % | Yatış ÷ poliklinik muayenesi × 100 |
| Hasta başına gelir | Gelir ÷ (poliklinik + yatış) |
| Hedef gerçekleşme % | Gerçekleşen ÷ hedef × 100 (poliklinik ve gelir; hedefi yazılmış günler üzerinden) |

Oranlar günlük oranların ortalaması değil, toplamların oranıdır.

**Uyarılar:** doluluk `--yuksek` (varsayılan %95) üstü · `--dusuk` (varsayılan %50) altı (bilgi) · poliklinik ya da gelir
hedef gerçekleşmesi `--hedef` (varsayılan %90) altı · dolu yatak yatak sayısından fazla.

**Örnek:** Dahiliye 14 günde 420 yatak-günün 400'ünde dolu → doluluk **%95,24** (eşik %95 → uyarı). Kardiyoloji
poliklinik hedefinin **%86,21**'ini gerçekleştirmiş.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                   # örnek veri (iki hafta, 6 bölüm)
python main.py --veriler bolum_gunluk_verileri.xlsx
python main.py --veriler bolum_gunluk_verileri.xlsx --yuksek 90 --dusuk 60 --hedef 95
```

Her satır bir bölümün bir günüdür. Zorunlu sütunlar: Tarih ve Bölüm; diğer göstergeler boş bırakılabilir.

## Çıktı

`cikti/hastane_operasyon_panosu.xlsx`:

- `Bölümler`: gelire göre sıralı bölüm toplamları ve hastane toplamı; doluluk ve hedefler renkli.
- `Günler`: hastane geneli günlük toplamlar.
- `Günlük Pano`: bölüm × gün göstergeleri.
- `Uyarılar`.

## Dikkat

- **Tanımlar kurum içinde tutarlı olmalı.** Poliklinik sayısına acil ve kontrol muayenelerinin girip girmediği, gelirin
  faturalanan mı tahsil edilen mi olduğu ve yatak sayısının aktif mi ruhsatlı mı olduğu bölümler arasında aynı
  olmalıdır.
- **Gün sonu sayımı:** Dolu yatak, günün belirli bir saatindeki (ör. gece yarısı) sayımdır; gün içi devir dahil
  değildir.
- **Eşikler politika varsayılanıdır.** Doluluk ve hedef sınırları bölümün yapısına göre (yoğun bakım, doğum, cerrahi)
  farklı olmalıdır.
- **Örnek veri:** Bölümler, sayılar, gelirler ve hedefler kurgusaldır.

## Testler

- Doluluk, poliklinikten yatış oranı, hasta başına gelir, hedef gerçekleşmesi, gelir payı, günlük toplamlar.
- Eşik uyarıları, dolu yatağın yatak sayısını aşması, yinelenen gün, okuma ve komut satırı.

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
