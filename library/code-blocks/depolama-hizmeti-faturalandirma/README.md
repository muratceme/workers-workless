# Depolama Hizmeti Faturalandırma · Kod Bloğu

> Depo ve Antrepo Yönetimi › Depo Şefi · Workers / Workless

Müşteri bazında **palet-gün, giriş-çıkış elleçleme ve katma değerli hizmetlerden aylık depolama faturasının hesabını**
çıkarır. İnternete bağlanmaz; fatura kesmez.

## Nasıl hesaplar?

| Kalem | Formül |
|---|---|
| Gün sonu stoku | Devir + o güne kadarki girişler − çıkışlar |
| Palet-gün | Dönemin her günü için gün sonu stokunun toplamı (giriş günü sayılır, çıkış günü sayılmaz) |
| Depolama bedeli | Palet-gün × palet-gün fiyatı |
| Elleçleme | Dönem içi giriş paleti × giriş fiyatı + çıkış paleti × çıkış fiyatı |
| Hizmetler | Adet × birim fiyat |
| Asgari fatura | Ara toplam asgarinin altındaysa fark eklenir |
| KDV, toplam | Ara toplam × `--kdv` % (varsayılan 20); ara toplam + KDV |

- Dönemden önceki hareketler ve "Devir" satırları açılış stokunu oluşturur; devir için elleçleme hesaplanmaz.
- `--cikis-gunu-dahil` verilirse çıkış yapılan gün de palet-güne sayılır.
- Fiyat önce müşteriye özel tarife satırından, yoksa müşterisi boş (genel) satırdan alınır.

**Örnek (Eylül, 30 gün):** devir 100 palet; 10.09'da 50 giriş; 20.09'da 30 çıkış → 9 × 100 + 10 × 150 + 11 × 120 = 3.720
palet-gün × 5,00 = 18.600; giriş 50 × 40 = 2.000; çıkış 30 × 40 = 1.200; etiketleme 500 × 1,50 = 750 → ara toplam 22.550;
KDV 4.510; toplam 27.060.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (dönem 2026-09)
python main.py --hareketler stok_hareketleri.xlsx --tarife tarife.xlsx --hizmetler hizmetler.xlsx --donem 2026-09
```

| Dosya | Sütunlar |
|---|---|
| Stok hareketleri | Müşteri, Tarih, Tür (Giriş / Çıkış / Devir), Palet |
| Tarife | Kalem (Depolama / Giriş / Çıkış / Asgari), Birim Fiyat (zorunlu); Müşteri (boş = genel) |
| Hizmetler (isteğe bağlı) | Müşteri, Hizmet, Adet, Birim Fiyat (zorunlu); Tarih |

## Çıktı

`cikti/depolama_faturasi.xlsx`:

- `Fatura Özeti`: müşteri başına devir, giriş, çıkış, dönem sonu stok, palet-gün, ara toplam, KDV, fatura toplamı.
- `Fatura Kalemleri`: her müşterinin fatura satırları (asgari fatura farkı vurgulu).
- `Stok Hareketi`: hareket olan günlerde gün sonu stoku (dökümü müşteriyle paylaşmak için).
- `Uyarılar`: eksiye düşen stok, fiyatı olmayan kalemler, okunamayan satırlar.

## Dikkat

- **Sözleşme esastır.** Palet-günün nasıl sayıldığı (gün sonu stoku, günün en yüksek stoku, giriş ve çıkış gününün
  dahil olması), asgari faturanın neyi kapsadığı ve dönem tanımı sözleşmeye göre değişir. Varsayılanlar bir tercihtir.
- **Palet esaslıdır.** Metrekare, raf gözü, koli veya ağırlık esaslı depolama ile kademeli fiyat desteklenmez.
- **KDV oranını siz doğrulayın.** Varsayılan %20 parametredir; antrepo, serbest bölge ve ihraç kayıtlı işlemlerdeki
  istisna ve özel durumları kod bilmez.
- **Devir stoku doğru olmalı.** Açılış stoku yanlışsa bütün ay yanlış çıkar; önceki dönemlerin hareketlerini dosyada
  bırakabilir veya dönem başından önceki tarihli bir "Devir" satırı yazabilirsiniz.
- **Tek para birimi.** Kur çevirisi yapılmaz.
- **Örnek veri:** Müşteriler, hareketler ve fiyatlar kurgusaldır.

## Testler

- Palet-gün (devir, ay içi giriş ve çıkış, son gün girişi, çıkış günü dahil, önceki ayların hareketleri).
- Müşteriye özel fiyat, hizmetler, asgari fatura farkı, KDV, eksi stok ve fiyatı olmayan kalem.
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
