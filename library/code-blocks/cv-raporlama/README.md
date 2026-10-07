# CV Raporlama · Kod Bloğu

> İnsan Kaynakları › İK Uzmanı · Workers / Workless

Bir klasördeki CV'leri okur ve tek bir Excel raporu üretir. **Kural tabanlıdır, internete
bağlanmaz, API anahtarı gerekmez.** Veriniz bilgisayarınızdan çıkmaz.

## Ne çıkarır?

| Alan | Nasıl |
|---|---|
| Ad soyad | İlk satırlardaki 2–4 kelimelik isim satırı (bulunamazsa dosya adı) |
| E-posta, telefon, LinkedIn | Düzenli ifadeler; telefon `+90 5xx xxx xx xx` biçimine çevrilir |
| Toplam deneyim (yıl) | `2018 - 2022`, `03/2019 - Günümüz` gibi aralıklar; çakışanlar birleştirilir, eğitim bölümü sayılmaz |
| En yüksek eğitim | Doktora › Yüksek Lisans › Lisans › Ön Lisans › Lise |
| Yabancı diller | Türkçe ve İngilizce dil adları |
| Beceriler | `beceriler.txt` listesindeki ifadeler |
| İlan uyum % | İlan verilirse: ilanda geçen becerilerden kaçı CV'de var |

## Kurulum

Python 3.10 veya üstü gerekir.

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
# Önce örnek veriyle deneyin
python main.py --ilan ornek_veri/ilan.txt

# Kendi CV'leriniz
python main.py --girdi "C:/İK/Başvurular" --cikti cikti/aday_raporu.xlsx --ilan ilan.txt
```

| Parametre | Varsayılan | Açıklama |
|---|---|---|
| `--girdi` | `ornek_veri/cvler` | CV klasörü (`.pdf`, `.docx`, `.txt`) |
| `--cikti` | `cikti/aday_raporu.xlsx` | Excel çıktısı |
| `--beceriler` | `beceriler.txt` | Aranacak beceriler, her satıra bir tane |
| `--ilan` | yok | İş ilanı metni (`.txt`); verilirse uyum yüzdesi hesaplanır |

## Çıktı

- **Adaylar** sayfası: her CV bir satır, ilan uyumuna göre sıralı, filtreli.
- **Özet** sayfası: toplam, okunan, ortalama deneyim, eğitim dağılımı.
- Okunamayan dosyalar raporu durdurmaz; **Uyarılar** sütununda belirtilir.

## Sınırlamalar

- Taranmış (resim) PDF'lerden metin çıkmaz → önce OCR uygulayın.
- Kural tabanlıdır: alışılmadık CV düzenlerinde alanlar boş kalabilir. Serbest metinde
  yorum gerekiyorsa bu görevin **AI Agent** sürümünü kullanın.
- Sonuçlar karar desteğidir; aday elemeden önce insan kontrolü şarttır.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
