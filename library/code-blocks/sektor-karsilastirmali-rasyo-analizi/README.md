# Sektör Karşılaştırmalı Rasyo Analizi · Kod Bloğu

> Krediler Tahsis › Kredi Tahsis Uzmanı · Workers / Workless

Firmanın bilanço ve gelir tablosu kalemlerinden **18 rasyoyu** hesaplar, sizin girdiğiniz **sektör değerleriyle**
karşılaştırır ve sapmaları hazır kalıp cümlelerle listeler. İnternete bağlanmaz; sektör verisi indirmez.

## Nasıl hesaplar?

| Grup | Rasyo | Formül | Olumlu yön |
|---|---|---|---|
| Likidite | Cari Oran | Dönen varlıklar ÷ KVYK | Yüksek |
| Likidite | Asit-Test Oranı | (Dönen varlıklar − stoklar) ÷ KVYK | Yüksek |
| Likidite | Nakit Oranı | Hazır değerler ÷ KVYK | Yüksek |
| Mali Yapı | Toplam Borç / Toplam Aktif (%) | (KVYK + UVYK) ÷ toplam varlıklar | Düşük |
| Mali Yapı | Özkaynak / Toplam Aktif (%) | Özkaynaklar ÷ toplam varlıklar | Yüksek |
| Mali Yapı | Borç / Özkaynak | (KVYK + UVYK) ÷ özkaynaklar | Düşük |
| Mali Yapı | Kısa Vadeli Borç / Toplam Borç (%) | KVYK ÷ (KVYK + UVYK) | Düşük |
| Kârlılık | Brüt Kâr Marjı (%) | (Net satışlar − satışların maliyeti) ÷ net satışlar | Yüksek |
| Kârlılık | Faaliyet Kâr Marjı (%) | Faaliyet kârı ÷ net satışlar | Yüksek |
| Kârlılık | Net Kâr Marjı (%) | Dönem net kârı ÷ net satışlar | Yüksek |
| Kârlılık | Aktif Kârlılığı (%) | Dönem net kârı ÷ toplam varlıklar | Yüksek |
| Kârlılık | Özkaynak Kârlılığı (%) | Dönem net kârı ÷ özkaynaklar | Yüksek |
| Faaliyet | Alacak Tahsil Süresi (gün) | Ticari alacaklar ÷ net satışlar × 365 | Düşük |
| Faaliyet | Stok Tutma Süresi (gün) | Stoklar ÷ satışların maliyeti × 365 | Düşük |
| Faaliyet | Ticari Borç Ödeme Süresi (gün) | Ticari borçlar ÷ satışların maliyeti × 365 | Yön verilmez |
| Faaliyet | Nakit Dönüş Süresi (gün) | Alacak süresi + stok süresi − borç ödeme süresi | Düşük |
| Faaliyet | Aktif Devir Hızı | Net satışlar ÷ toplam varlıklar | Yüksek |
| Borç Servisi | Faiz Karşılama Oranı | Faaliyet kârı ÷ finansman giderleri | Yüksek |

KVYK: kısa vadeli yükümlülükler; UVYK: uzun vadeli yükümlülükler. "Olumlu yön" kredi değerlendirmesi açısındandır.

- **Sapma %** = (firma − sektör) ÷ |sektör| × 100. Mutlak sapma `--tolerans` (varsayılan %10) içindeyse "Sektöre yakın".
- **Çeyrek konumu:** alt ve üst çeyrek girdiyseniz firma değeri bu aralığa göre konumlanır.
- **Önceki döneme göre:** birden çok dönem varsa son iki dönem karşılaştırılır (İyileşti / Kötüleşti).
- **Dikkat notları:** özkaynak ≤ 0, faiz karşılama < 1, dönem zararı, cari oran < 1.
- Özkaynak sıfır veya negatifse borç / özkaynak ve özkaynak kârlılığı hesaplanmaz.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                        # örnek veri
python main.py --firma firma_finansallari.xlsx --sektor sektor_degerleri.xlsx --tolerans 15
```

**Firma dosyası:** ilk sütun `Kalem`, sonraki sütunlar dönemler (ör. `2024`, `2025`). Sektörle karşılaştırılan dönem
son sütundur. Tanınan kalem adları:

Dönen Varlıklar, Hazır Değerler (Nakit ve Nakit Benzerleri), Ticari Alacaklar, Stoklar, Toplam Varlıklar (Aktif
Toplamı), Kısa Vadeli Yükümlülükler (KVYK), Uzun Vadeli Yükümlülükler (UVYK), Ticari Borçlar, Özkaynaklar, Net
Satışlar (Hasılat), Satışların Maliyeti, Faaliyet Kârı (Esas Faaliyet Kârı), Finansman Giderleri, Dönem Net Kârı.

Satışların maliyeti ve finansman giderleri eksi işaretli girilmişse mutlak değeri alınır.

**Sektör dosyası:** `Rasyo`, `Sektör Ortalaması` (veya `Medyan`), isteğe bağlı `Alt Çeyrek`, `Üst Çeyrek`. Rasyo
adlarını yukarıdaki tablodaki gibi yazın. Girmediğiniz rasyolar "Sektör verisi yok" olarak kalır.

## Çıktı

`cikti/rasyo_karsilastirma.xlsx`:

- `Karşılaştırma`: firma, sektör, sapma, durum, çeyrek konumu, önceki döneme göre yön.
- `Dönemsel Rasyolar`: tüm dönemler için rasyolar.
- `Yorumlar`: dikkat notları ve sapma cümleleri; boş "Analist Notu" sütunu.
- `Uyarılar`: eksik kalem, okunamayan tutar, bilanço denkliği, tanınmayan sektör rasyosu.

## Dikkat

- **Sektör değerlerini siz girersiniz.** Pakette gömülü sektör verisi yoktur. Kaynak olarak örneğin TCMB Sektör
  Bilançoları kullanılabilir; değerin kaynağını, yılını ve sektör kodunu kendi notunuza yazın.
- **Tanım farkı sonucu değiştirir.** Kaynağınızın rasyo formülü bu tablodakinden farklıysa (ör. ortalama bakiye,
  FAVÖK bazlı faiz karşılama, 360 gün) karşılaştırma yanıltıcı olur. Formülleri eşleştirin.
- **Süreler dönem sonu bakiyesiyle** ve 365 gün üzerinden hesaplanır; ortalama bakiye kullanılmaz. Ara dönem
  (ör. 6 aylık) gelir tablosuyla süre ve kârlılık rasyoları yıllık değerlerle karşılaştırılamaz.
- **Yorumlar kalıptır.** Cümleler sapmanın yönünü söyler, nedenini açıklamaz. Kredi kararı veya derecelendirme değildir.
- **Enflasyon düzeltmesi:** Firma ve sektör verisinin aynı esasla (düzeltilmiş / düzeltilmemiş) hazırlandığını kontrol edin.
- **Örnek veri:** Firma finansalları ve sektör değerleri kurgusaldır; gerçek bir sektörün verisi değildir.

## Testler

- 18 rasyonun elle hesaplanmış değerleri; eksik kalem ve negatif özkaynak.
- Sapma, tolerans sınırı, olumlu / olumsuz yön, çeyrek konumu, yön verilmeyen rasyo.
- Dikkat notları ve dönemsel yön.
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
