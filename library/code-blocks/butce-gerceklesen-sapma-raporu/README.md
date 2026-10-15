# Bütçe-Gerçekleşen Sapma Raporu · Kod Bloğu

> Finans › Bütçe ve Raporlama Uzmanı · Workers / Workless

Yıllık bütçeyi muhasebeden alınan **gerçekleşen kayıtlarla** hesap ve masraf merkezi bazında karşılaştırır.
Rapor ayı ve yılbaşından bugüne (YTD) sapmayı TL ve % olarak, **lehte / aleyhte** yönüyle verir; önemlilik
eşiğini aşan sapmaları işaretler ve yöneticinin dolduracağı bir **açıklama sütunu** bırakır. İnternete bağlanmaz.

## Neler yapar?

- **Hesap eşleşmesi:** Muavindeki ayrıntılı hesaplar (`770.02.001`) bütçedeki satıra (`770.02`) en uzun önek
  eşleşmesiyle toplanır. Bütçede karşılığı olmayan hesaplar **Bütçelenmemiş** olarak ayrıca listelenir.
- **Gelir / gider yönü:** `Tür` sütunu yoksa Tekdüzen Hesap Planı'na göre belirlenir: 60, 64, 67 grupları gelir;
  61-63, 65, 66, 68, 69 ve 7'li maliyet hesapları gider. Gelirde fazla gerçekleşme lehte, giderde aleyhtedir.
- **Borç / alacak:** Muavin borç-alacak sütunlarıyla verilirse gider = borç − alacak, gelir = alacak − borç
  (iptal ve iadeler kendiliğinden düşülür).
- **Önemli sapma:** |sapma| ≥ `--esik-tutar` TL **ve** ≥ bütçenin `--esik-yuzde`'si (varsayılan 10.000 TL ve %10).
- **Zamanlama uyarısı:** Ay sapması önemli ama YTD sapması önemsizse "dönem kayması olabilir" notu düşülür.
- **Yıl sonu tahmini:** YTD gerçekleşen + kalan ayların bütçesi.
- Gerçekleşende birden çok yıl varsa en son yıl alınır (`--yil` ile değiştirilebilir).

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                  # örnek bütçe ve muavinle dener (Eylül 2026)
python main.py --butce butce_2026.xlsx --gerceklesen muavin.xlsx --ay 9
python main.py --butce butce.xlsx --gerceklesen muavin.xlsx --esik-yuzde 5 --esik-tutar 25000
```

**Bütçe:** geniş biçim (Hesap Kodu, Hesap Adı, Masraf Merkezi, Ocak … Aralık) ya da uzun biçim (Hesap Kodu, Ay, Tutar).
**Gerçekleşen:** muavin / yevmiye dökümü (Tarih, Hesap Kodu, Masraf Merkezi, Borç, Alacak) ya da Ay + Tutar.
Masraf merkezi yalnız bir dosyada varsa karşılaştırma hesap bazında yapılır.

## Çıktı

`Özet` · `Sapma Raporu` (filtreli, açıklama sütunlu) · `Masraf Merkezi` (grafikli) · `Aylık Seyir` (grafikli) · `Bilgi`

## Testler

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
