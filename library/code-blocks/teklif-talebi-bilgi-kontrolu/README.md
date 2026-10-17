# Teklif Talebi Bilgi Kontrolü · Kod Bloğu

> Teknik (Risk Kabul / Underwriting) › Teknik Uzman Yardımcısı · Gerisi Bende

Acenteden gelen teklif taleplerini **sizin hazırladığınız zorunlu alan kurallarına** göre kontrol eder; eksik, geçersiz
ve tutarsız bilgileri talep ve acente bazında listeler, acenteye gönderilecek eksik metnini üretir. İnternete bağlanmaz.

## Nasıl hesaplar?

Kural dosyasındaki her satır bir zorunlu bilgidir:

| Sütun | Anlamı |
|---|---|
| Branş | Kuralın uygulanacağı branş; `*` veya boş = tüm branşlar; birden çok branş `Yangın\|İşyeri` |
| Alan | Talep dosyasındaki sütun adı (büyük-küçük harf ve Türkçe karakter ayrımı yapılmaz) |
| Tür | `Metin`, `Sayı`, `Tarih` veya `Seçenek` (boşsa Metin) |
| Seçenekler | Tür `Seçenek` ise geçerli değerler: `Var\|Yok` |
| Asgari | Metin için en az karakter sayısı; Sayı için en küçük değer |
| Koşul | `Alan=Değer`: kural yalnız o alan o değerdeyse uygulanır (ör. `Önceki Hasar=Var`) |
| Açıklama | Eksik listesinde acenteye gösterilecek not |

| Durum | Ne zaman |
|---|---|
| Eksik | Alan boş, `-` veya `?` |
| Geçersiz | Metin asgari uzunluğun altında; sayı okunamıyor, sıfır / negatif veya asgarinin altında; tarih okunamıyor; değer seçeneklerde yok |
| Tutarsız | Bitiş tarihi başlangıçtan sonra değil; `Toplam Sigorta Bedeli` diğer `... Bedeli` sütunlarının toplamından 1 TL'den fazla farklı |

Tutarlılık kontrolleri yalnız ilgili sütunlar dosyada varsa çalışır. Hiç eksiği olmayan talep "Tam" sayılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --talepler talepler.xlsx --kurallar kurallar.xlsx
```

| Dosya | Sütunlar |
|---|---|
| Talepler | Talep No, Branş (zorunlu), Acente, Sigortalı + kurallarda adı geçen sütunlar |
| Kurallar | Branş, Alan, Tür, Seçenekler, Asgari, Koşul, Açıklama |

## Çıktı

`cikti/teklif_talebi_kontrol.xlsx`:

- `Talep Özeti`: talep bazında sonuç ve eksik alanlar.
- `Eksikler`: her eksik için durum ve açıklama; boş "Tamamlandı" sütunu.
- `Acente Özeti`: acente bazında eksikli talep oranı ve en sık eksik alanlar.
- `Alan Özeti`: hangi bilginin kaç talepte eksik olduğu.
- `Uyarılar`: okunamayan kurallar, dosyada olmayan kural alanları, mükerrer talepler.

`cikti/acente_eksik_notlari.txt`: acente bazında, talep talep eksik listesi (e-postaya yapıştırılabilir).

## Dikkat

- **Zorunlu alan listesi sizindir.** Örnek kural dosyası yalnız biçimi gösterir; bir sigorta şirketinin risk kabul
  yönetmeliği veya mevzuat listesi değildir. Yapı tarzı seçenekleri gibi değerleri kendi sınıflandırmanıza göre yazın.
- **Doluluk kontrolüdür, doğruluk kontrolü değildir.** Kod adresin gerçekten var olduğunu, bedelin yeterli olduğunu,
  beyan edilen faaliyetin doğru olduğunu veya hasar geçmişinin tam bildirildiğini denetlemez.
- **Risk kabul kararı değildir.** "Tam" yalnız kuralların aradığı alanların dolu olduğunu gösterir.
- **Kişisel veri:** Talep dosyaları kimlik numarası ve adres içerir; KVKK ve şirket içi kurallara göre saklayın.
- **Örnek veri:** Acenteler, sigortalılar, adresler ve kimlik numaraları kurgusaldır.

## Testler

- Dört kural türü: geçerli, geçersiz ve boş değerler; asgari sınırlar.
- Branş eşleşmesi, koşullu kural, kuralı olmayan branş.
- Tarih ve bedel tutarlılığı; acente ve alan özetleri; acente metni.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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
