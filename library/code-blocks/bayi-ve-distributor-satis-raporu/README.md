# Bayi ve Distribütör Satış Raporu · Kod Bloğu

> Satış ve Kanal Yönetimi › Kanal Satış Yöneticisi · Gerisi Bende

Bayi bazında satış, hedef, tahsilat, alacak ve stok verisinden **hedef gerçekleşme, tahsilat oranı, alacak günü ve
stok gününü** hesaplar; sorunlu bayileri işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Hedef gerçekleşme % | Satış ÷ hedef × 100 |
| Büyüme % | (Satış − önceki dönem satışı) ÷ önceki dönem satışı × 100 |
| Tahsilat oranı % | Dönem tahsilatı ÷ dönem satışı × 100 |
| Alacak gün | Açık alacak ÷ satış × dönem gün sayısı |
| Stok gün | Bayi stoku ÷ bayinin kendi satışı × dönem gün sayısı |
| Satış payı % | Bayi satışı ÷ toplam satış × 100 |

"Bayi Satışı" (bayinin kendi müşterisine satışı) verilmemişse stok gün, bayiye yapılan satışa göre hesaplanır.

İşaretler: gerçekleşme `--hedef-esik` (%80) altında · alacak gün `--alacak-gun` (60) üstünde · stok gün `--stok-gun`
(45) üstünde · dönem satışı sıfır ama alacak veya stok var.

**Örnek:** Satış 800.000, hedef 1.000.000, açık alacak 1.900.000, bayi stoku 1.500.000, bayi satışı 700.000; dönem
30 gün. Gerçekleşme %80; alacak gün 1.900.000 ÷ 800.000 × 30 = 71,25; stok gün 1.500.000 ÷ 700.000 × 30 = 64,29 →
"Alacak günü yüksek", "Stok günü yüksek".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --bayiler bayiler.xlsx
python main.py --bayiler bayiler.xlsx --gun 90 --hedef-esik 85 --alacak-gun 75 --stok-gun 30
```

Zorunlu sütunlar: Bayi, Satış. Diğerleri: Bölge, Hedef, Tahsilat, Açık Alacak, Bayi Stoku, Bayi Satışı, Önceki Dönem
Satış. Boş bırakılan alanın göstergesi hesaplanmaz.

## Çıktı

`cikti/bayi_satis_raporu.xlsx`:

- `Bayiler`: satışa göre sıralı; tüm göstergeler, işaretler ve aksiyon için boş sütun.
- `Bölge Özeti`: bölge ve toplam bazında aynı göstergeler.
- `Uyarılar`: işaretlenen bayiler gerekçesiyle; okunamayan satırlar.

## Dikkat

- **Tutarlar aynı esasta olmalı.** Hepsi KDV dahil veya hepsi hariç. Stok ve bayi satışı aynı fiyat düzeyinde (ikisi de
  bayi alış fiyatıyla veya ikisi de satış fiyatıyla) değilse stok gün yanıltır.
- **Alacak gün dönem satışına göre kabadır.** Vadesi gelmemiş alacak ile gecikmiş alacağı ayırmaz; yaşlandırma için
  "Alacak Yaşlandırma" paketini kullanın.
- **Tahsilat oranı %100'ü aşabilir.** Dönem içinde önceki dönemlerin alacağı tahsil edildiyse normaldir.
- **Eşikler sektöre göre değişir.** Hızlı tüketimde 45 gün stok yüksek, dayanıklı tüketimde normal olabilir;
  eşikleri kendi kanal politikanıza göre verin.
- **Bayiye satış, bayinin satışı değildir.** Bayiye satış hedefi tutarken bayi stoku birikiyorsa sonraki dönem satışı
  düşer; stok günü bu yüzden izlenir.
- **Örnek veri:** Bayiler, bölgeler ve tutarlar kurgusaldır.

## Testler

- Tüm göstergelerin hesabı, boş alanlar, bayi satışı olmadan stok gün.
- Her işaretin eşik sınırı; satışı sıfır bayiler; parametreyle değişen eşikler.
- Bölge ve toplam özetleri; okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

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
