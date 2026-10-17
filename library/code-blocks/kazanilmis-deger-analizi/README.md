# Kazanılmış Değer Analizi · Kod Bloğu

> Planlama ve Proje Kontrol › Planlama Mühendisi · Gerisi Bende

Planlanan değer, kazanılmış değer ve gerçekleşen maliyetten **SPI, CPI ve tamamlanma maliyet tahminini** hesaplar
(kazanılmış değer yöntemi, EVM). İnternete bağlanmaz.

## Nasıl hesaplar?

PV = planlanan değer, EV = kazanılmış değer, AC = gerçekleşen maliyet (kümülatif), BAC = toplam bütçe.

| Gösterge | Formül |
|---|---|
| Program sapması (SV) | EV − PV |
| Maliyet sapması (CV) | EV − AC |
| Program performansı (SPI) | EV ÷ PV |
| Maliyet performansı (CPI) | EV ÷ AC |
| Tamamlanma tahmini (EAC) | BAC ÷ CPI |
| EAC, kalan iş bütçesinde biterse | AC + (BAC − EV) |
| EAC, iki verim birlikte sürerse | AC + (BAC − EV) ÷ (CPI × SPI) |
| Kalan maliyet (ETC) | EAC − AC |
| Tamamlanma sapması (VAC) | BAC − EAC |
| Gereken verim (TCPI) | (BAC − EV) ÷ (BAC − AC) |
| Süre tahmini (kaba) | Planlanan dönem sayısı ÷ SPI |

- Veriler **dönemsel** kabul edilir ve kod toplar; dosyanız kümülatifse `--kumulatif` verin.
- `--bac` verilmezse BAC, planlanan değerlerin toplamıdır.
- EV ve AC'si boş dönemler ileri dönemdir; analiz, verisi olan son döneme göre yapılır.

**Örnek:** BAC 12.000.000; üçüncü dönem sonunda PV 6.000.000, EV 5.100.000, AC 6.000.000 → SV −900.000, CV −900.000,
SPI 0,85, CPI 0,85; EAC = 12.000.000 × 6.000.000 ÷ 5.100.000 = 14.117.647,06; VAC −2.117.647,06; TCPI = 6.900.000 ÷
6.000.000 = 1,15.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --donemler donemler.xlsx --bac 12000000
python main.py --donemler kumulatif_donemler.xlsx --kumulatif
```

Sütunlar: Dönem, Planlanan Değer (zorunlu); Kazanılmış Değer, Gerçekleşen Maliyet (ileri dönemlerde boş).

## Çıktı

`cikti/kazanilmis_deger.xlsx`:

- `Özet`: analiz dönemindeki tüm göstergeler, açıklamalarıyla.
- `Dönem Bazında`: kümülatif PV, EV, AC ve dönem sonu SV, CV, SPI, CPI.
- `Uyarılar`: eksik dönem verisi, BAC ile plan toplamı farkı, bütçeyi aşan kazanılmış değer.

## Dikkat

- **Kazanılmış değer ölçüme dayanmalı.** EV, tamamlanan işin *bütçe fiyatıyla* değeridir; harcanan paradan veya geçen
  süreden türetilirse CPI ve SPI anlamsızlaşır. Tamamlanma yüzdesini metraj veya hakedişten alın.
- **SPI proje sonuna doğru yanıltır.** İş bittiğinde EV, PV'ye eşitlenir ve SPI gecikmiş projede bile 1'e yaklaşır.
  Süre tahmini kabadır; gecikme için kritik yol analizi yapın.
- **Üç EAC, üç varsayımdır.** Hangi tahminin gerçekçi olduğu sapmanın nedenine bağlıdır (tek seferlik mi, kalıcı mı).
- **Aynı esasta veri.** PV, EV ve AC aynı kapsamda (KDV, genel gider, fiyat farkı dahil veya hariç) ve aynı para
  biriminde olmalıdır. AC'ye taahhüt edilmiş ama faturalanmamış maliyetler eklenmelidir.
- **İş kalemi kırılımı yoktur.** Kalem bazında sapma için "Proje Maliyet Kontrol Raporu" paketine bakın.
- **Örnek veri:** Dönemler ve tutarlar kurgusaldır.

## Testler

- SV, CV, SPI, CPI, üç EAC, ETC, VAC, TCPI ve süre tahmini; dönem bazında kümülatif değerler.
- Kümülatif girdi, elle verilen BAC, programın ilerisinde ve bütçenin altında durum, EV veya AC'si 0 olan dönem.
- Boş dönem, azalan kümülatif değer ve okuma hataları; örnek veri uçtan uca.

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
