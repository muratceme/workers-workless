# Sınav Sonuç ve Madde Analizi · Kod Bloğu

> Eğitim › Ölçme Değerlendirme ve Rehberlik › Ölçme Değerlendirme Uzmanı · Workers / Workless

Optik okuyucu dökümünden veya cevap tablosundan öğrenci **netlerini ve puanlarını** hesaplar. Klasik test
kuramına göre **madde analizi** (güçlük, ayırt edicilik, çeldirici) ve **KR-20 güvenirlik** hesaplar. Kazanım
bazında sınıf başarısını raporlar. İnternete bağlanmaz.

## Ne hesaplar?

| Ölçüt | Tanım |
|---|---|
| Net | doğru − yanlış ÷ k (`--yanlis-katsayi`, varsayılan 4; 4 seçenekli sınavlarda 3; yanlış götürmüyorsa 0) |
| Puan | net ÷ değerlendirilen soru sayısı × 100 |
| Güçlük (p) | doğru cevaplayanların oranı |
| Ayırt edicilik (d) | üst %27 grubun p'si − alt %27 grubun p'si |
| Madde-toplam r | madde puanı ile o madde çıkarılmış toplam puan arasındaki korelasyon |
| KR-20 | k/(k−1) × (1 − Σ p(1−p) / σ²) |
| Ölçmenin standart hatası | σ × √(1 − KR-20) |

**Yorumlar**
- **Ayırt edicilik (Ebel ölçütleri):**
  - ≥ 0,40 çok iyi
  - 0,30–0,39 iyi
  - 0,20–0,29 düzeltilmeli
  - < 0,20 çıkarılmalı veya yeniden yazılmalı
- **Güçlük:**
  - < 0,20 çok zor
  - 0,20–0,39 zor
  - 0,40–0,60 orta
  - 0,61–0,80 kolay
  - > 0,80 çok kolay
- **Çeldirici analizi:** Yanlış bir şıkkı üst grup alt gruptan belirgin biçimde daha çok seçiyorsa uyarı verir;
  anahtar hatası veya belirsiz soru olabilir. Hiç seçilmeyen çeldiriciler de işaretlenir.

**Kitapçıklar ve iptal**
- **Farklı soru sırası:** B kitapçığı farklı soru sırasındaysa anahtar dosyasındaki `A Soru No` sütunuyla ortak
  numaraya çevrilir. Böylece madde analizi tüm kitapçıklar üzerinden yapılır.
- **İptal edilen sorular:** Varsayılan olarak değerlendirme dışı bırakılır. `--iptal-dogru` ile herkese doğru
  sayılır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                        # örnek: 40 öğrenci, 2 şube, A/B kitapçık
python main.py --cevaplar cevaplar.xlsx --anahtar anahtar.xlsx
python main.py --cevaplar optik.csv --anahtar anahtar.csv --yanlis-katsayi 3 --iptal-dogru
```

**Cevap dosyası**
- **Sütunlar:** Öğrenci No, Ad Soyad, Sınıf, Kitapçık ve soru sütunları (1, 2, 3 … veya S1, S2 …).
- **Tek sütunlu döküm:** Optik okuyucu cevapları tek bir dizi olarak veriyorsa `Cevaplar` sütunu kullanılabilir
  (ör. `ABCD AB...`; boşluk = boş).
- **Okunamayan işaretler:** Çoklu işaretleme veya okunamayan işaret (`*`) yanlış sayılır.

**Anahtar dosyası**
- **Uzun biçim:** Kitapçık, Soru, Cevap, [A Soru No, Kazanım, İptal].
- **Geniş biçim:** Kitapçık, 1, 2, 3 … (her kitapçık bir satır).

## Çıktı

`Özet` (KR-20, sınıf karşılaştırması) · `Öğrenci Sonuçları` (genel ve sınıf sırası) · `Madde Analizi` (renkli
ayırt edicilik, çeldirici notları, şık dağılımları) · `Kazanım Analizi` (genel ve sınıf bazında başarı yüzdesi) ·
`Bilgi`

## Dikkat

- **Küçük gruplar:** Madde istatistikleri küçük gruplarda (ör. 30'un altında öğrenci) dalgalı olabilir. Kararları
  birden fazla uygulamanın sonuçlarına dayandırın.
- **Öğrenci verisi:** Öğrenci sonuçları kişisel veridir. Raporu yetkisiz kişilerle paylaşmayın.

## Testler

Testler elle hesaplanmış küçük bir sınavla p, d, KR-20 ve net değerlerini doğrular.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
