# Kıdem ve İhbar Tazminatı Hesaplama · Kod Bloğu

> İnsan Kaynakları › Bordro ve Özlük İşleri Uzmanı · Gerisi Bende

İşe giriş-çıkış tarihi ve ücretten **kıdem tazminatını** ve **ihbar tazminatını** hesaplar; vergi
kesintileriyle net tutarları çıkarır. Tek kişi veya toplu çıkış listesi için çalışır. İnternete bağlanmaz.

## Kurallar

| Kalem | Kural |
|---|---|
| Hizmet süresi | Giriş ve çıkış günleri dahil yıl / ay / gün |
| Giydirilmiş ücret | Çıplak brüt + düzenli aylık yan ödemeler (yemek, yol vb.) + yıllık ikramiye ÷ 12 |
| Kıdem | Her tam yıl için 30 günlük giydirilmiş ücret; artan ay ÷ 12, gün ÷ 365 oranında (1475 s. Kanun md. 14). En az 1 yıl hizmet gerekir |
| Kıdem tavanı | Fesih tarihindeki tavan: 01.01-30.06.2026 → 64.948,77 TL · 01.07-31.12.2026 → 73.729,87 TL (2025 dönemleri de tanımlı) |
| Kıdem kesintisi | Gelir vergisinden istisna (GVK md. 25/7); yalnız damga vergisi (binde 7,59) |
| İhbar süresi | 6 aydan az 2 hafta · 6 ay-1,5 yıl 4 hafta · 1,5-3 yıl 6 hafta · 3 yıldan fazla 8 hafta (4857 s. Kanun md. 17) |
| İhbar tutarı | Giydirilmiş günlük ücret (aylık ÷ 30) × gün; tavan uygulanmaz |
| İhbar kesintisi | Gelir vergisi (çıkış yılındaki kümülatif matraha göre dilim) + damga vergisi |

Parametreler `tr_parametreler.json` dosyasındadır ve her biri kaynaklıdır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py --giris 15.03.2019 --cikis 30.09.2026 --brut 60000 --yan-odeme 4500 --ikramiye 60000 --kumulatif 450000
python main.py --girdi cikislar.xlsx
python main.py                     # örnek listeyle dener
```

Toplu girdi sütunları: **Ad Soyad, İşe Giriş, İşten Çıkış, Brüt** (zorunlu); Yan Ödeme, Yıllık İkramiye,
Kümülatif Matrah, Kıdem Hakkı (E/H), İhbar Hakkı (E/H) (isteğe bağlı). Tarihler `GG.AA.YYYY` veya `YYYY-AA-GG`.

## Önemli

Kıdem ve ihbar hakkının doğup doğmadığı **fesih nedenine** bağlıdır (raporun "Kurallar" sayfasında özetlenmiştir).
Bu araç hak durumunu belirlemez; "Kıdem Hakkı / İhbar Hakkı" sütunlarıyla siz belirtirsiniz. Kullanılmayan
yıllık izin ücreti dahil değildir. Sonuçlar bilgilendirme amaçlıdır; uyuşmazlıkta hukuki destek alın.

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
| 1475 sayılı İş Kanunu (kıdem tazminatı) | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.1475.pdf) | Resmî metinle karşılaştırılan maddeler: 14 |
| 193 sayılı Gelir Vergisi Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.4.193.pdf) | Resmî metinle karşılaştırılan maddeler: 25, 94 |

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
