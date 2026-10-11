# Kapasite Darboğaz Analizi · Kod Bloğu

> Üretim Planlama › Üretim Planlama Uzmanı · Workers / Workless

İş merkezi × hafta bazında **yükleme oranını** hesaplar; kapasiteyi aşan haftaları (darboğaz) bulur ve iş emirlerini
boş haftalara **kaydırma önerisi** çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Yük (saat) = "Yük (saat)" sütunu;  boşsa (adet × birim süre dk + hazırlık dk) ÷ 60
Yükleme %  = haftanın toplam yükü ÷ haftalık kapasite × 100
```

| Durum | Koşul |
|---|---|
| Darboğaz | Yük > kapasite (tam %100 darboğaz değildir) |
| Sınırda | Yükleme ≥ `--uyari` (varsayılan %90) |
| Uygun | Altı |

**Kaydırma önerisi** her darboğaz haftası için şöyle üretilir:

1. Aşımı tek başına kapatan en küçük iş emri önce denenir; yoksa en büyük iş emri.
2. İş emri, tamamının sığdığı en yakın haftaya taşınır (1 hafta önce, 1 hafta sonra, 2 hafta önce ...). Eşit
   uzaklıkta öne çekme tercih edilir. Hedef hafta kapasitesini aşacaksa o hafta seçilmez.
3. Termin haftasını aşmayan seçenek varsa o seçilir; yoksa öneri "termin riski" ile işaretlenir.
4. Aşım sürüyorsa aynı adımlar yinelenir. Hiçbir iş emri hiçbir haftaya sığmıyorsa aşım "çözülemedi" olarak kalır
   (fazla mesai, fason veya iş emrini bölme gerekir).

Planlama ufku, dosyadaki ilk ve son hafta arasıdır; ufkun dışına kaydırma önerilmez. Daha önce bir darboğazı
çözmek için taşınan iş emri ikinci kez taşınmaz.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --is-emirleri is_emirleri.xlsx --kapasiteler kapasiteler.xlsx --uyari 85
```

| Dosya | Sütunlar |
|---|---|
| İş emirleri | İş Emri, İş Merkezi, Hafta (`2026-H42` veya o haftadan bir tarih), Yük (saat) veya Adet + Birim Süre (dk) + Hazırlık (dk), Termin Haftası |
| Kapasiteler | İş Merkezi, Haftalık Kapasite (saat), Hafta (boş = her hafta; dolu = yalnız o hafta, ör. tatil haftası) |

Haftalar ISO hafta numarasıdır (pazartesi başlar).

## Çıktı

`cikti/kapasite_darbogaz.xlsx`:

- `Yükleme Tablosu`: iş merkezi × hafta; yük, kapasite, yükleme %, durum ve öneri sonrası durum.
- `Darboğazlar`: aşım saatleri ve kaydırmayla çözülüp çözülmediği.
- `Kaydırma Önerileri`: iş emri, mevcut ve önerilen hafta, yön, termin riski; boş "Karar" sütunu.
- `Uyarılar`: kapasitesi tanımsız iş merkezleri, okunamayan satırlar.

## Dikkat

- **Yalnız kapasiteye bakar.** Operasyon sırası (rota), önceki ve sonraki operasyonların haftası, malzeme ve kalıp
  hazırlığı, personel ve vardiya düzeni kontrol edilmez. Öne çekilen bir operasyonun önceki operasyonu henüz
  bitmemiş olabilir.
- **Öneri en iyi çözüm değildir.** Basit bir kuralla üretilir; çizelgeleme optimizasyonu yapılmaz.
- **Kapasite sizin tanımınızdır.** Vardiya sayısı, verimlilik ve planlı bakım düşülmüş net saat girin.
- **İş emri bölünmez.** Kısmi kaydırma önerilmez; bölme kararı planlamacıya aittir.
- **Örnek veri:** İş merkezleri, iş emirleri ve kapasiteler kurgusaldır.

## Testler

- Hafta yazımları, yükleme oranı ve durum sınırları (tam kapasite darboğaz değildir).
- Öne çekme, aşımı kapatan en küçük iş emri, termin riski olmayan seçeneğin tercihi.
- Çözülemeyen aşım, birden çok kaydırma, haftaya özel kapasite, kapasitesi tanımsız iş merkezi.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
