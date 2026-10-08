# Kredi Komitesi Değerlendirme Notu · AI Agent

> Bankacılık › Krediler Tahsis › Kredi Tahsis Uzmanı · Workers / Workless

Teklif bilgisi, limit ve teminat tabloları, rasyolar ve istihbarat/analiz notlarından **kredi komitesine sunulacak
değerlendirme notunun taslağını** hazırlar.
- **Kod:** Limit, teminat karşılama ve rasyo hesaplarını ve kontrollerini yapar; tablolar koddan gelir.
- **Model:** Güçlü ve zayıf yönleri, riskleri ve azaltıcı unsurları, öneriyi, şartları ve izleme kriterlerini yazar.

Öneri taslaktır; **kredi kararı yetkili komiteye aittir.**

> **Önemli — bankacılık sırrı:** Kredi dosyası bankacılık sırrı ve ticari sır içerir. Dış bir yapay zekâ servisine
> göndermek bankanızın politikasına aykırı olabilir; bilgi güvenliği ve uyum birimlerinizin onayını alın. Veri
> bilgisayardan çıkmasın isteniyorsa `.env` içinde `WW_PROVIDER=ollama` ile yerel model kullanın.

## Nasıl çalışır?

1. **Dosyaları tanır (kod):** adında "teklif", "limit", "teminat", "rasyo", "istihbarat", "analiz" geçen dosyalar;
   tablolar `.csv` / `.xlsx`, metinler `.txt` / `.pdf` / `.docx`.
2. **Hesaplar ve kontrol eder (kod):**

   | Kontrol | Önem |
   |---|---|
   | Teklifte talep gerekçesi veya ödeme kaynağı yok | orta |
   | Toplam limit artışı %50 ve üzeri | orta |
   | Yeni kredi türü; limit doluluğu ≥ %90 | bilgi |
   | Risk mevcut limiti aşıyor | yüksek |
   | Teminatlar talebi karşılamıyor (tesis edilecekler dahil) | orta |
   | Mevcut teminatlar yetersiz, tesis edilecek teminat şarta bağlanmalı | orta |
   | Rasyolarda kötüleşme: %10 – %25 | bilgi |
   | Rasyolarda kötüleşme: %25 ve üzeri | orta |
   | Borç servis karşılama (DSCR) < 1 | yüksek |
   | DSCR 1 – 1,2 | orta |
   | Cari oran < 1 | orta |
   | Faiz karşılama < 1; özkaynak negatif | yüksek |
   | Bankacılık Kanunu md. 54: risk grubu toplamı banka özkaynağının %25'ini aşıyor (`--banka-ozkaynak` verilirse) | yüksek |
   | Metinlerde "gecikme", "icra", "protesto" gibi ifadeler ("bulunmamaktadır" gibi olumsuzlananlar sayılmaz) | orta / yüksek |

   Rasyonun iyi yönü adından anlaşılır (ör. cari oran yükselirse, kaldıraç ve devir süreleri düşerse iyi).
3. **Maskeler:**
   - Firma unvanı `[FİRMA]`, VKN `[VKN]`, risk grubu `[GRUP]` olur.
   - Müşteri temsilcisi ve `--gizle` ile verilen ortak/kefil adları `[KİŞİ-n]` olur; soyadları da maskelenir.
   - Telefon, e-posta, TCKN ve IBAN maskelenir.
   - Veri gönderilmeden önce onayınız alınır.
4. **Yazar (model):** talep özeti, firma ve faaliyet, mali değerlendirme, teminat değerlendirmesi, güçlü/zayıf
   yönler, riskler ve azaltıcı unsurlar, öneri (Olumlu / Şartlı olumlu / Ek bilgi gerekli / Olumsuz), şartlar
   (kullandırım öncesi / sonrası / sürekli), izleme kriterleri ve eksik bilgiler.
5. **Modelin çıktısını denetler (kod):**
   - Kodda yüksek önemli bulgu varken öneri "Olumlu" ise uyarır.
   - Tesis edilecek bir teminat şartlarda yer almıyorsa uyarır.
   - Metindeki, girdilerde olmayan sayıları işaretler.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                        # örnek: kurgusal gıda üreticisi
python agent.py --girdi ./teklif_klasoru --gizle "Ortak Adı" "Kefil Adı"
python agent.py --girdi ./teklif --banka-ozkaynak 25000000000          # BK md. 54 kontrolüyle
```

| Dosya | Biçim |
|---|---|
| `teklif.txt` | `Firma:`, `VKN:`, `Sektör:`, `Risk Grubu:`, `Grup Riski (Bankamız):` (bu firma hariç), `Müşteri Temsilcisi:`, `Talep Gerekçesi:`, `Ödeme Kaynağı:` |
| `limitler.csv` | Kredi Türü, Mevcut Limit, Mevcut Risk, Talep Edilen Limit, Vade, Fiyat ("mektup", "akreditif", "aval" geçenler gayrinakdi sayılır) |
| `teminatlar.csv` | Teminat Türü, Değer, Durum (Mevcut / Tesis edilecek), Açıklama |
| `rasyolar.csv` | İlk sütun rasyo adı, sonraki sütunlar dönemler (ör. 2024, 2025); yüzdeler `%` ile |
| `istihbarat*.txt`, `analist_notu.txt`, `ziyaret*.txt` | Serbest metin |

## Çıktı

- **Not:** `komite_notu.md`, takma adlar geri açılmış; limit, rasyo ve teminat tabloları koddan.
- **Excel:** `komite_notu.xlsx`:
  - `Özet`: boş "Komite Kararı" ve "Karar Gerekçesi" satırları.
  - `Kontroller`: kod bulguları, model riskleri ve azaltıcı unsurlar, "İnceleme" sütunu.
  - `Limitler`, `Teminatlar`, `Rasyolar` (değişim ve kötüleşme işareti).
  - `Şartlar ve İzleme`: "Komite Kararı" sütunu.

## Dikkat

- **Taslak niteliği:** Not bir taslaktır. Rakamları kaynak belgelerle, şartları bankanızın kredi politikasıyla
  karşılaştırın.
- **Teminat değerleri:** Ham değerlerdir; bankanızın teminat katsayıları uygulanmaz.
- **Eşikler:** DSCR 1 – 1,2 ve rasyo kötüleşme eşikleri yaygın kullanılan genel ölçülerdir; kredi politikanızdaki
  eşikleri esas alın.
- **BK md. 54:** Kontrol talep toplamı ile girilen grup riskini basitçe toplar. Kredi sayılan işlemler, indirimler
  ve risk grubu tanımı için Kredi İşlemlerine İlişkin Yönetmelik hükümlerini esas alın.
- **Örnek veri:** Örnek firma, kişiler ve bankalar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Dosyaların tanınması ve teklif alanları.
- Limit ve teminat toplamları.
- Örnekteki bulgular: %155 limit artışı, %70,6 mevcut teminat karşılaması, rasyo kötüleşmeleri, DSCR 1,12,
  BK md. 54 aşımı.
- Kişi, firma ve grup adlarının modele gitmemesi.
- "Olumlu" öneri ile yüksek bulgu çelişkisi; şartlarda olmayan makine rehni.
- Uydurma bir sayının (45.000.000 TL) yakalanması.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka ve banka politikasına
uygunluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
