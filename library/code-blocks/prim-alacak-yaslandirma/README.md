# Prim Alacak Yaşlandırma · Kod Bloğu

> Tahsilat Yönetimi › Tahsilat Uzmanı · Gerisi Bende

Tahsil edilmemiş prim taksitlerini **vade gününe göre yaşlandırır**; acente, branş ve poliçe bazında raporlar ve
takip edilecek poliçeleri listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Kalan alacak = taksit tutarı − tahsil edilen        (0 veya altındaysa taksit kapalıdır, rapora girmez)
Gecikme günü = rapor tarihi − vade tarihi
```

| Aralık | Gecikme günü |
|---|---|
| Vadesi gelmemiş | Vade rapor günü veya sonrası (0 gün ve altı) |
| 1-30 | 1 – 30 gün |
| 31-60 | 31 – 60 gün |
| 61-90 | 61 – 90 gün |
| 90+ | 90 günden fazla |

- Sınırlar `--araliklar 30,60,90` ile değiştirilir (ör. `15,45`).
- **Vadesi geçmiş %** = vadesi geçmiş kalan alacak ÷ toplam açık alacak.
- **Poliçe bazında** gecikme günü, poliçenin en eski açık taksitine göredir.
- **Takip listesi:** en eski açık taksiti `--takip-gun` (varsayılan 30) günden **fazla** gecikmiş poliçeler.
- Kısmi tahsilatta yalnız kalan tutar yaşlandırılır. Bir taksitteki fazla tahsilat diğer taksitlere mahsup edilmez;
  uyarı olarak gösterilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 01.10.2026)
python main.py --taksitler prim_taksitleri.xlsx --tarih 31.12.2026 --araliklar 15,45,90 --takip-gun 15
```

Sütunlar: Poliçe No, Vade Tarihi, Taksit Tutarı (zorunlu); Acente, Sigortalı, Branş, Taksit No, Tahsil Edilen.
Her satır bir taksittir.

## Çıktı

`cikti/prim_alacak_yaslandirma.xlsx`:

- `Yaşlandırma`: aralık bazında kalan alacak ve pay; vadesi geçmiş toplam.
- `Acente Bazında`, `Branş Bazında`: aralıklara dağılım, vadesi geçmiş oran, en uzun gecikme.
- `Poliçe Bazında`: açık ve gecikmiş taksit sayısı, en eski vade, gecikme günü.
- `Takip Listesi`: aranacak poliçeler; boş "Yapılan İşlem" ve "Söz Verilen Tarih" sütunları.
- `Açık Taksitler`: taksit bazında ayrıntı.
- `Uyarılar`: okunamayan satırlar, fazla tahsilatlar.

## Dikkat

- **Hukuki süreç değerlendirmesi yapmaz.** Prim borcunda ihtar, temerrüt, teminatın durması ve poliçenin feshi
  koşulları Türk Ticaret Kanunu, ilgili genel şartlar ve poliçe hükümlerine tabidir; takip günü eşiği yalnız iç iş
  takibi içindir.
- **Tahsilat kaydı günceldir varsayılır.** Kredi kartı blokeleri, yoldaki havaleler ve acente cari hesabındaki
  mahsuplar dosyaya yansımadıysa alacak olduğundan yüksek görünür.
- **Sorumlu taraf:** Kod, primi acentenin mi sigortalının mı ödeyeceğini ayırmaz; acente bazındaki tablo poliçenin
  acentesine göredir.
- **İptal ve zeyiller:** İptal edilen poliçelerin iade / tenzil kayıtlarını dosyadan çıkarın veya taksit tutarına yansıtın.
- **Karşılık hesabı değildir.** Şüpheli alacak karşılığı veya mevzuat raporlaması üretmez.
- **Örnek veri:** Acenteler, sigortalılar ve tutarlar kurgusaldır.

## Testler

- Aralık sınırları (0, 1, 30, 31, 60, 61, 90, 91 gün) ve özel aralıklar.
- Genel, acente ve branş dağılımı; vadesi geçmiş oran.
- Poliçe bazında en eski vade, takip listesi sınırı (tam 30 gün girmez).
- Kısmi ve fazla tahsilat, okuma hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6102 sayılı Türk Ticaret Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6102.pdf) | Resmî metinle karşılaştırılan maddeler: 5/A, 22, 796, 1434 |

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
