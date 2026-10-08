# İhale Dokümanı Özeti · AI Agent

> İnşaat › İhale ve Teklif › İhale Mühendisi · Workers / Workless

İhale dokümanından **yeterlik kriterlerini, teminatları, süreleri, cezaları, istenen belgeleri ve özel
şartları** özetler. Kamu (4734 sayılı Kanun) ve özel sektör ihalelerinde kullanılabilir.
- **Model:** Her kalemi **kaynak madde ve kelimesi kelimesine alıntıyla** yazar.
- **Kod:** Alıntıların belgede gerçekten geçtiğini kontrol eder, belgeler arası çelişkileri bulur ve tutarları
  hesaplar.

## Nasıl çalışır?

1. **Okur ve tarar (kod):**
   - Klasördeki belgeleri okur (`.pdf` metin tabanlı, `.docx`, `.txt`): idari şartname, sözleşme tasarısı,
     teknik şartname, ilan, zeyilname.
   - **İşin süresi** ve **günlük gecikme cezası** oranı belgeler arasında farklıysa "Hata" verir.
   - İhale tarihini bulur.
   - Belgelerde geçen önemli konuları listeler: teminat, iş deneyimi, bilanço, fiyat farkı, avans, sigorta,
     iş artışı…
2. **Maskeler:** Telefon, e-posta ve `--gizle` ile verilen adlar maskelenir. Veri gönderilmeden önce onayınız
   alınır.
3. **Özetler (model):** İhale bilgilerini, yeterlik kriterlerini, teminatları, süreleri, cezaları, istenen
   belgeleri ve özel şartları çıkarır. Özel şartlara yüksek/orta/düşük risk verir. Belirsizlik ve çelişkiler için
   **açıklama talebi** soruları önerir.
4. **Doğrular (kod):**
   - **Alıntı:** Her kalemin alıntısı belgede aranır; bulunamayan "Alıntı belgede bulunamadı" olarak
     işaretlenir.
   - **Oran:** Oranın (ör. %50, binde 0,6) alıntıda rakamla geçip geçmediği kontrol edilir.
   - **Atlanan konu:** Belgede geçen ama özette hiç yer almayan konular (ör. sigorta) listelenir.
5. **Hesaplar (kod):** `--teklif` verilirse şunları hesaplar:
   - Geçici ve kesin teminat tutarı.
   - İş deneyimi ve iş hacmi (ciro, taahhüt) alt sınırları.
   - Günlük gecikme cezası.
   - Takvim: ihaleye ve açıklama talebi son gününe kalan gün.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                    # örnek: kurgusal pazar yeri yapım ihalesi
python agent.py --girdi ./ihale_dokumani --teklif 48500000
python agent.py --girdi idari_sartname.pdf --teklif 48.500.000 --yaklasik 52000000 --bugun 08.10.2026
```

- Belge türü dosya adından (`idari`, `sozlesme`, `teknik`, `ilan`, `zeyil`) veya ilk satırlardan tanınır.
- Taranmış (görüntü) PDF'ler okunmaz; önce OCR uygulayın.

## Çıktı

`ihale_ozeti.xlsx`:
- `Özet`: ihale bilgileri (kaynak ve doğrulamayla), takvim, genel özet.
- `Yeterlik Kriterleri`: "Firmamız Karşılıyor mu?" sütunuyla.
- `Teminat ve Cezalar`.
- `Hesaplar`: oran × taban = tutar.
- `Süreler`.
- `Özel Şartlar ve Riskler`.
- `İstenen Belgeler`: "Hazır mı?" ve "Sorumlu" sütunlarıyla.
- `Kontroller`: çelişkiler, doğrulanamayan alıntılar, özette olmayan konular.
- `Açıklama Talepleri`.

Ayrıca `ihale_ozeti.md` özet.

## Dikkat

- **Taslak niteliği:** Çıktı taslaktır. Teklif kararından önce dokümanın tamamı (zeyilnameler dahil) ve ilgili
  mevzuat okunmalıdır: kamu yapım işlerinde Kamu İhale Kanunu, Yapım İşleri İhaleleri Uygulama Yönetmeliği ve
  Kamu İhale Kurumu'nun güncel düzenlemeleri.
- **Mevzuata uygunluk:** Model mevzuattan değer eklemez; yalnız belgede yazanı özetler. Belgedeki bir şartın
  mevzuata uygun olup olmadığını değerlendirmez.
- **Tutarlar:** Hesaplanan tutarlar alt sınırdır. İş deneyim belgelerinin güncellenmesi (endeksleme), benzer iş
  grubu ve belge tutarının nasıl hesaplanacağı için şartnameye ve mevzuata bakın.
- **Süreler:** Açıklama talebi son günü, şartnamedeki ifadeye göre basit gün çıkarmasıyla hesaplanır. Gün
  sayımı ve tatil kurallarını kontrol edin.
- **Örnek veri:** Örnek ihale belgeleri tamamen kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Örnekte bilerek bırakılmış süre çelişkisi: idari şartnamede 300, sözleşmede 330 takvim günü.
- Uydurma bir alıntının yakalanması.
- Özette atlanan "sigorta" konusunun bulunması.
- Teminat, iş deneyimi, ciro ve gecikme cezası hesapları.
- Takvim.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin uygunluğu kullanıcının
sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
