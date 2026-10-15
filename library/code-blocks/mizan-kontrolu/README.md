# Mizan Kontrolü · Kod Bloğu

> Muhasebe › Genel Muhasebe Uzmanı · Workers / Workless

Aylık mizanı **Tekdüzen Hesap Planı** kurallarına göre kontrol eder. Ters bakiyeli hesapları, mantık dışı
bakiyeleri ve önceki aya göre olağandışı değişimleri açıklama ve öneriyle listeler. İnternete bağlanmaz.

## Kontroller

**Biçim (Hata)**
- **Denklik:** Borç toplamı alacak toplamına, borç bakiyeler alacak bakiyelere eşit mi?
- **Satır aritmetiği:** Her satırda borç − alacak = bakiye tutuyor mu? Aynı satırda hem borç hem alacak bakiyesi
  var mı?
- **Alt hesap toplamı:** Ana hesap, alt hesaplarının toplamına eşit mi? (120 = 120.01 + 120.02 …)

**Ters bakiye**

| Hesap | Kural |
|---|---|
| 1-2 grubu | Borç bakiye |
| 3-4-5 grubu | Alacak bakiye |
| Adı "(-)" ile biten düzenleyici hesaplar | Ters yönde bakiye (103, 122, 129, 257, 322, 371, 501, 580, 591 …) |
| 60, 64, 67 | Alacak bakiye |
| 61-63, 65, 66, 68 | Borç bakiye |
| 7/A | x0 gider hesabı borç, x1 yansıtma hesabı alacak bakiye |

Ana hesap normal görünse de ters bakiyeli **alt hesaplar** ayrıca listelenir ve virman önerilir:

| Alt hesap durumu | Öneri |
|---|---|
| Müşteri hesabı alacak bakiyede | 340 Alınan Sipariş Avansları'na virman |
| Satıcı hesabı borç bakiyede | 159 Verilen Sipariş Avansları'na virman |
| Banka hesabı alacak bakiyede (KMH) | 300 Banka Kredileri'ne virman |
| Personele borçlar borç bakiyede | 196 Personel Avansları'na virman |

**Mantık**
- **Kasa şişkinliği:** Kasa bakiyesi aktif toplamının `--kasa-orani` (varsayılan %3) üzerindeyse uyarı verir.
- **131/231 Ortaklardan alacaklar:** Adat hesaplanması ve faiz faturası düzenlenmesi önerilir (KVK md. 13).
- **KDV mahsubu:** 191 ve 391 birlikte bakiye veriyorsa mahsup kaydı eksik olabilir.
- **7/A yansıtma:** Yansıtma hesabı gider hesabını kapatmıyorsa uyarı verir.
- **Mantık dışı tutarlar:**
  - Birikmiş amortisman maddi duran varlıklardan fazla.
  - Şüpheli alacak karşılığı alacaktan fazla.
  - 590 ve 591 birlikte bakiye veriyor.

**Önceki aya göre** (`--onceki`)
- **Olağandışı değişim:** Değişim hem `--esik` (100.000 TL) hem `--esik-oran` (%50) eşiğini aşıyorsa uyarı verir.
- **Hesap durumu değişimi:** Yön değiştiren, yeni bakiye veren ve kapanan hesaplar listelenir.
- **Kümülatif azalma:** 6 ve 7 grubu hesaplar yıl başından itibaren birikir. Önceki aya göre azalma, iptal veya ters
  kayıt belirtisidir. Yıl başı (satışlar yarıdan fazla azalmışsa) otomatik algılanır ve bu kontrol atlanır.

Bulgular beş seviyede verilir: **Hata · Yüksek · Orta · Dikkat · Bilgi**. Bulgular sayfasında açıklama ve
sorumlu için boş bir sütun bulunur.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                               # örnek Eylül/Ağustos mizanlarıyla
python main.py --mizan eylul.xlsx --onceki agustos.xlsx
python main.py --mizan eylul.xlsx --esik 50000 --esik-oran 30 --kasa-orani 2
```

**Girdi biçimi**
- **Kaynak:** Muhasebe programından alınan mizan dökümü kullanılır. Üstteki firma ve tarih satırları atlanır.
- **Sütunlar:** `Hesap Kodu` ile birlikte şu bakiye sütunlarından biri gerekir:
  - `Borç Bakiye` / `Alacak Bakiye`
  - `Borç` / `Alacak`
  - tek bir `Bakiye` sütunu (borç +, alacak −)
- **Hesap kırılımları:** Noktalı (120.01.001) ve noktasız (12001001) kodlar desteklenir.
- **Ana hesap satırı yoksa:** Ana hesap bakiyesi alt kırılımlardan toplanır.

## Çıktı

`Özet` · `Bulgular` · `Ana Hesaplar` (önceki ay ve değişim dahil) · `Bilgi`

## Dikkat

Bu araç kontrol önerisi üretir. Virman ve düzeltme kayıtları ile vergi değerlendirmeleri (adat, kasa, KDV) için
mali müşavirinize danışın.

## Testler

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 5520 sayılı Kurumlar Vergisi Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.5520.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

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
