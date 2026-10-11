# Teklif Talebi Bilgi Kontrolü · Kod Bloğu

> Teknik (Risk Kabul / Underwriting) › Teknik Uzman Yardımcısı · Workers / Workless

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

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
