# Yıllık İzin Hakedişi Hesaplama · Kod Bloğu

> İnsan Kaynakları › Bordro ve Özlük İşleri Uzmanı · Gerisi Bende

Çalışanların yıllık ücretli izin hakkını **4857 sayılı İş Kanunu md. 53**'e göre hizmet yılı bazında
hesaplar: toplam hakediş, kullanılan, kalan izin, sonraki hakediş tarihi ve gün sayısı. Brüt ücret
verilirse kullanılmayan iznin brüt ücret karşılığını da tahmin eder. İnternete bağlanmaz.

## Kurallar

| Hakediş anındaki hizmet yılı | İzin |
|---|---|
| 1 – 5 yıl (5 dahil) | 14 gün |
| 5 yıldan fazla, 15 yıldan az | 20 gün |
| 15 yıl ve üzeri | 26 gün |

- Hakediş tarihinde **18 yaş ve altı** veya **50 yaş ve üstü** olanlara en az 20 gün.
- **Yer altı** işlerinde çalışanlara +4 gün.
- Hak, her hizmet yılı tamamlandığında doğar; yaş kuralı her hakediş tarihindeki yaşa göre uygulanır.
- 29 Şubat'ta işe girenlerin hakedişi artık olmayan yıllarda 28 Şubat'ta doğar.

Toplu iş sözleşmesi veya iş sözleşmesi daha uzun izin öngörebilir; araç kanuni asgariyi hesaplar.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py --girdi personel.xlsx                 # bugünün tarihine göre
python main.py --girdi personel.xlsx --tarih 31.12.2026
python main.py --tarih 30.09.2026                     # örnek listeyle dener
```

| Sütun | Zorunlu | Açıklama |
|---|---|---|
| Ad Soyad | Evet | |
| İşe Giriş | Evet | `GG.AA.YYYY` veya `YYYY-AA-GG` |
| Doğum Tarihi | Hayır | Yaş kuralı için gerekli |
| Kullanılan İzin | Hayır | Bugüne kadar kullanılan toplam gün |
| Devreden İzin | Hayır | Sisteme geçiş öncesinden devreden gün |
| Yer Altı | Hayır | E/H |
| Brüt | Hayır | Kullanılmayan izin ücreti tahmini için (brüt ÷ 30 × kalan gün) |

## Çıktı

- **İzin Durumu:** kişi bazında özet; kullanılan izin hakedişi aşıyorsa satır kırmızı.
- **Hakediş Detayı:** her hizmet yılı için hakediş tarihi, o tarihteki yaş ve gün.
- **Kurallar:** referans tarih ve uygulanan kurallar.

## Testler

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |

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
