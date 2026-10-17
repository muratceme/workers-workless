# Tedavi Fiyat Teklifi Hazırlama · Kod Bloğu

> Uluslararası Hasta Hizmetleri › Uluslararası Hasta Koordinatörü · Gerisi Bende

Tedavi planı ve paket fiyat listesinden **konaklama ve transfer dahil hastaya gönderilecek teklif dosyasını** hazırlar:
Excel teklif tablosu ve Türkçe veya İngilizce teklif metni. İnternete bağlanmaz. Hasta adı gerekmez.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Kalem eşleştirme | Plandaki kalem, fiyat listesinde kod veya ad ile aranır (büyük-küçük harf ayrımı yapılmaz) |
| Kalem tutarı | Adet × birim fiyat (adet boşsa 1) |
| Ara toplam | Kategori bazında: Tedavi, Konaklama, Transfer, Diğer |
| İndirim | Toplam × `--indirim` %; `--indirim-kategori` verilirse yalnız o kategorinin ara toplamına |
| Genel toplam | Toplam − indirim |
| Geçerlilik | Teklif tarihi + `--gecerlilik-gun` (varsayılan 30) |

- Fiyat listesinde bulunamayan kalem teklife alınmaz ve yüksek önemli uyarı verilir.
- Tüm kalemler aynı para biriminde olmalıdır; farklıysa kod durur. Kur çevirisi yapılmaz.
- `--dil en` ile kalem adları fiyat listesindeki "İngilizce Ad" sütunundan alınır; boşsa Türkçe adı yazılır ve uyarılır.

**Örnek:** 4 × 600 + 8 × 180 + 1 × 30 = 3.870 (tedavi), 5 gece × 70 = 350 (konaklama), 2 × 40 = 80 (transfer) → toplam
4.300; %5 indirim 215 → genel toplam 4.085 EUR.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (12.10.2026, %5 indirim)
python main.py --plan tedavi_plani.xlsx --fiyatlar fiyat_listesi.xlsx --hasta-ref UH-2026-014 --dil en
python main.py --plan p.xlsx --fiyatlar f.xlsx --indirim 10 --indirim-kategori tedavi --gecerlilik-gun 15 --not "Uçak bileti dahil değildir."
```

| Dosya | Sütunlar |
|---|---|
| Tedavi planı | Kalem (fiyat listesindeki kod veya ad; zorunlu); Adet, Not |
| Fiyat listesi | Kalem, Birim Fiyat (zorunlu); Kod, Kategori, Para Birimi, İngilizce Ad |

## Çıktı

- `cikti/tedavi_teklifi.xlsx` → `Teklif` (kategori bazında kalemler, ara toplamlar, indirim, genel toplam) ve `Uyarılar`.
- `cikti/tedavi_teklifi.txt` → e-postaya yapıştırılabilir teklif metni.

Örnek metin:

```
TEDAVİ FİYAT TEKLİFİ
Referans: UH-2026-001 · Teklif tarihi: 12.10.2026 · Geçerlilik: 11.11.2026

Tedavi
  - Diş implantı (tek) · 4 × 600,00 EUR = 2.400,00 EUR
  ...
Toplam: 4.300,00 EUR
İndirim (%5): -215,00 EUR
GENEL TOPLAM: 4.085,00 EUR
```

## Dikkat

- **Teklif taslaktır, onaysız gönderilmez.** Tedavi planını hekim belirler; kod plandaki kalemlerin tıbbi uygunluğunu
  değerlendirmez. Teklif hastaya gitmeden önce hekim ve kurum onayından geçmelidir.
- **Mevzuat ve tanıtım kuralları sizin sorumluluğunuzdadır.** Uluslararası sağlık turizmi yetki belgesi, fiyat
  bildirimi, sağlık hizmetlerinde tanıtım ve bilgilendirme kuralları ile sonuç vaadi yasağı gibi yükümlülükleri kod
  bilmez. Metne eklediğiniz notlarda garanti veya kesin sonuç ifadesi kullanmayın.
- **Vergi hesaplanmaz.** Fiyatların KDV dahil mi hariç mi olduğu fiyat listenize bağlıdır; kod KDV veya istisna
  hesabı yapmaz. Gerekirse `--not` ile belirtin ve mali müşavirinize danışın.
- **Fiyat listeniz güncel olmalı.** Kod fiyat bilmez; listede ne yazıyorsa onu kullanır.
- **Eksik kalem teklifi eksiltir.** Fiyat listesinde bulunamayan kalem toplama girmez; `Uyarılar` sayfasını
  kontrol etmeden göndermeyin.
- **Kişisel sağlık verisi:** Hasta adı, pasaport numarası ve tanı bilgisi dosyalara girilmez; teklifte yalnız
  `--hasta-ref` ile verdiğiniz referans kodu görünür.
- **Örnek veri:** Kalemler ve fiyatlar kurgusaldır; gerçek fiyat göstergesi değildir.

## Testler

- Kalem tutarı, kategori ara toplamları, toplam ve kategoriye özel indirim, geçerlilik tarihi.
- Bulunamayan kalem, farklı para birimi ve teklifte olmayan kategoriye indirim hataları.
- Türkçe ve İngilizce teklif metni, sayı biçimi, ek notlar.
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
