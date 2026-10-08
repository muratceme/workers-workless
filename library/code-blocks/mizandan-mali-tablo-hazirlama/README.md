# Mizandan Mali Tablo Hazırlama · Kod Bloğu

> Muhasebe › Muhasebe Müdürü · Workers / Workless

Tekdüzen Hesap Planı'na göre tutulan **mizandan**, Muhasebe Sistemi Uygulama Genel Tebliği (MSUGT) biçiminde
**ayrıntılı bilanço** ve **gelir tablosu** hazırlar. Önceki dönem mizanı verilirse karşılaştırmalı sütun ekler.
İnternete bağlanmaz.

## Ne yapar?

**Bilanço**
- **Düzen:** I. Dönen Varlıklar, II. Duran Varlıklar, III. Kısa Vadeli Yabancı Kaynaklar, IV. Uzun Vadeli Yabancı
  Kaynaklar ve V. Özkaynaklar bölümlerinden oluşur.
- **Gruplar ve hesaplar:** Her bölüm grup başlıklarına (A. Hazır Değerler, B. Menkul Kıymetler …) ve 3 haneli ana
  hesaplara ayrılır.
- **Düzenleyici hesaplar:** Birikmiş amortismanlar ve şüpheli alacak karşılığı gibi (-) hesaplar eksi tutarla
  gösterilir.

**Gelir tablosu**
- **Düzen:** A. Brüt Satışlar'dan başlayıp ara toplamlarla K. Vergi Karşılıkları ve Dönem Net Kârı'na kadar
  ilerler.
- **Dönem kârı:** Gelir tablosu hesapları açıksa (kapanış öncesi mizan) dönem kârı hesaplanır ve özkaynaklarda
  590 veya 591 olarak gösterilir. Kapanış sonrası mizanda kâr 590/591'den alınır.

**Ters bakiye virmanları**

Alt kırılımda ters bakiye veren cari ve banka hesapları, bilançoda karşı tarafta gösterilir. Kapatmak için
`--virman-yok` kullanın.

| Ters bakiyeli alt hesap | Bilançoda |
|---|---|
| 102 Bankalar (alacak bakiye, KMH) | 300 Banka Kredileri |
| 120/121/127 Alıcılar (alacak bakiye) | 340 Alınan Sipariş Avansları |
| 320/321/329 Satıcılar (borç bakiye) | 159 Verilen Sipariş Avansları |
| 335 Personele Borçlar (borç bakiye) | 196 Personel Avansları |
| 220/221 ve 420/421 (uzun vadeli) | 440 / 259 |

**Kontroller**
- Aktif ve pasif toplamlarının eşitliği.
- 7'li maliyet hesaplarında kalan bakiye (yansıtma kaydı yapılmamış olabilir).
- 690/692/697/698 hesaplarında bakiye (kapanış kaydı tamamlanmamış olabilir).
- Hem gelir tablosu hesaplarının açık olması hem 590/591'de bakiye bulunması (kâr çift sayılabilir).

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                     # örnek 2026/2025 mizanlarıyla
python main.py --mizan mizan_2026.xlsx --onceki mizan_2025.xlsx --unvan "Örnek A.Ş."
python main.py --mizan mizan_2026.xlsx --virman-yok
```

Mizan dosyası muhasebe programından alınan döküm olabilir:
- **Başlık satırları:** Üstteki firma ve tarih satırları atlanır.
- **Sütunlar:** `Hesap Kodu` ile birlikte `Borç Bakiye`/`Alacak Bakiye`, `Borç`/`Alacak` ya da tek bir `Bakiye`
  sütunu gerekir.
- **Kırılımlar:** Ana hesap satırı yoksa alt kırılımlar toplanır ve ana hesap adı Tekdüzen listesinden verilir.

## Çıktı

`Bilgi` (denklik ve uyarılar) · `Bilanço Aktif` · `Bilanço Pasif` · `Gelir Tablosu` · `Virmanlar`

## Dikkat

- **Dönem sonu kayıtları:** Değerleme, karşılık, amortisman, reeskont ve kapanış kayıtları mizana işlenmiş
  olmalıdır.
- **Raporlama çerçevesi:** Tablolar VUK/MSUGT biçimindedir; TFRS veya BOBİ FRS raporlaması değildir.
- **Yasal kullanım:** Beyanname ve yasal defterler için mali müşavirinizin kontrolünden geçirin.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
