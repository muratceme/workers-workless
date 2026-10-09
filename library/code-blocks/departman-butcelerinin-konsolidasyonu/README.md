# Departman Bütçelerinin Konsolidasyonu · Kod Bloğu

> Finans › Bütçe ve Raporlama Uzmanı · Workers / Workless

Departmanlardan gelen bütçe şablonlarını (bir klasördeki Excel dosyaları) tek dosyada birleştirir. Format ve toplam
hatalarını dosya ve hücre adresiyle listeler, tavan ve önceki yıla göre kontrol eder. İnternete bağlanmaz.

## Şablon

Her departman bir `.xlsx` dosyası gönderir (`ornek_veri/departmanlar/` içindeki örnekler gibi):
- Üstte `Departman:` etiketi ve yanında departman adı. Yoksa dosya adı kullanılır. İsteğe bağlı `Hazırlayan:`.
- Başlık satırı ilk 20 satırda aranır: `Hesap Kodu`, `Gider Kalemi`, `Ocak` … `Aralık` (Oca, Şub gibi kısaltmalar da
  olur), `Toplam` (isteğe bağlı), `Açıklama`.
- Sayfa adı `Bütçe` ise o sayfa, değilse ilk sayfa okunur.

## Neleri kontrol eder?

| Kontrol | Önem |
|---|---|
| Başlık bulunamadı, dosya açılamadı, boş bütçe | Yüksek |
| Aynı departman için birden çok dosya (ada göre sonuncusu kullanılır) | Yüksek |
| Toplam sütunu ≠ ayların toplamı (metin sayı varsa Excel TOPLA uyarısıyla) | Yüksek |
| Sayı olmayan hücre; hesaplanmamış formül | Yüksek |
| Hesap planında olmayan kod | Yüksek |
| Bütçe tavanı aşımı | Yüksek |
| Önceki yıl gerçekleşene göre `--degisim` (%30) üstü değişim | Orta |
| Negatif tutar, aynı hesabın tekrarı (toplanır), hesap kodu boş satır, eksik ay sütunu | Orta |
| Boş ay (0 sayılır), metin olarak girilmiş sayı, hesap adı farklı | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 5 dosya (biri mükerrer), 4 departman
python main.py --klasor butceler/ --hesap-plani plan.xlsx --tavanlar tavanlar.xlsx --degisim 30
```

| Dosya | Sütunlar |
|---|---|
| Hesap planı (isteğe bağlı) | Hesap Kodu, Hesap Adı |
| Tavanlar (isteğe bağlı) | Departman, Bütçe Tavanı, Önceki Yıl Gerçekleşen |

## Çıktı

`konsolide_butce.xlsx`:
- `Hatalar`: önem, tür, departman, dosya, hücre adresi, açıklama; boş "Düzeltildi" sütunu. Departmanlara geri
  bildirim için kullanılır.
- `Konsolide Bütçe`: hesap × ay ve yıllık toplam.
- `Departman × Hesap`: yıllık tutarlar.
- `Departman Özeti`: dosya, hazırlayan, bütçe, tavan ve farkı, önceki yıl, değişim, yüksek önemli hata sayısı; boş
  "Onay" sütunu ve grafik.
- `Uzun Liste`: departman / hesap / ay / tutar; pivot tablo için.

## Dikkat

- **Formüller:** Formüllü şablonlar Excel'de açılıp kaydedilmiş olmalıdır. Aksi halde formül değeri okunamaz ve
  "Hesaplanmamış formül" hatası verilir.
- **Mükerrer dosya:** Aynı departmandan birden çok dosya geldiğinde dosya adına göre sıralanır ve sonuncusu
  kullanılır. Revize dosyaları adlandırırken bunu dikkate alın veya eski dosyayı klasörden çıkarın.
- **Örnek veri:** Örnek şablonlar ve tutarlar kurgusaldır.

## Testler

Şunlar test edilir:
- Departman ve konsolide toplamlar; mükerrer dosya seçimi.
- Tüm hata türleri: toplam, metin sayı, hesap planı, tekrar, negatif, boş ay, tavan, değişim.
- Departman adı olmayan şablon, formül, sayı olmayan hücre ve şablon olmayan dosya; Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
