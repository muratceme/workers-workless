# Ar-Ge Proje Raporu Taslağı · AI Agent

> Ar-Ge ve Ürün Geliştirme › Ar-Ge Mühendisi · Workers / Workless

İş paketi ve bütçe tabloları ile proje notlarından **Ar-Ge projesi dönem raporu taslağı** hazırlar (TÜBİTAK TEYDEB,
Ar-Ge / tasarım merkezi gibi programların dönem raporlarında istenen tipik başlıklar). İlerleme ve bütçe sapmalarını
**kod hesaplar**. Model yalnız notlara dayanarak yazar: notta olmayan sonuç, sayı ya da patent / yayın iddiası ve
ilerlemesi %100 olmayan pakete "tamamlandı" demek yakalanır.

## Nasıl çalışır?

1. **Hesap (kod):** Her iş paketi için gerçekleşen − planlanan ilerleme. −10 puan ve altı "Gecikmeli", +10 ve üstü
   "Planın önünde", %100 "Tamamlandı". Bütçe kalemlerinde sapma oranı; ±%10'u aşan kalemler açıklama gerektirir (`--esik`).
2. **Taslak (model):** Dönem özeti; her iş paketi için yapılanlar, sonuçlar, sapma açıklaması, sonraki dönem planı (not
   numaralarıyla); bütçe sapma açıklamaları; sorunlar ve çözümler; çıktılar.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Bir iş paketi ele alınmamış; gecikmeli pakette sapma açıklaması yok; bölüm not numarası göstermiyor
   - Paket bölümünde o paketin notlarında / tablosunda olmayan sayı (ör. "12 haftada üretildi")
   - İlerlemesi %100 olmayan paket için "İP2 başarıyla tamamlanmıştır"
   - Açıklama gereken bütçe kalemi açıklanmamış
   - Notlarda geçmeyen patent, faydalı model, yayın, lisans ya da ticarileşme iddiası
4. **Rapor (kod):** Künye, iş paketi ilerleme özeti, kalem bazında ve toplam harcama tablosu eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                           # örnek (kurgusal) kestirimci bakım projesi
python agent.py --ip is_paketleri.xlsx --butce butce.xlsx --notlar notlar.docx --esik 15
```

Not dosyası örneği:

```
Proje Adı: …
Program: TÜBİTAK TEYDEB 1501
Dönem: 01.04.2026 - 30.09.2026

Notlar:
- İP2: Sensör kartının ilk sürümü üretildi; gürültü nedeniyle kart revize ediliyor.
- Bütçe: Ek ölçüm cihazı alındığından alet-teçhizat kalemi aşıldı.
```

## Çıktı

- `cikti/arge_donem_raporu.txt`: rapor taslağı.
- `cikti/arge_donem_raporu.xlsx`: İş Paketleri (durum renkli), Bütçe (açıklama gereken kalemler sarı), Notlar.

## Dikkat

- **Format sizin.** Programların rapor formları ve harcama belgelendirme kuralları (ör. TÜBİTAK harcama listeleri, Ar-Ge
  merkezi faaliyet raporu) dönem dönem değişir. Taslağı kurumun güncel formuna aktarın; harcamaları mali müşavir / mali
  sorumlu ile doğrulayın.
- **Gizli bilgi.** Teknik notlar yapay zekâ sağlayıcısına gönderilir. Ticari sır ya da patent başvurusu öncesi buluş
  ayrıntısı içeriyorsa notları buna göre hazırlayın ya da yerel model (Ollama) kullanın.
- **Örnek veri:** Proje kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta %65 ilerlemedeki İP2 için "başarıyla tamamlanmıştır" der, notta
olmayan "12 hafta" yazar, sapma açıklamasını boş bırakır, İP4'ü ve seyahat sapmasını atlar, olmayan bir patent başvurusu
ekler. Kodun altı sorunu da yakaladığı; sapma, durum ve bütçe hesaplarının doğru olduğu doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve raporun doğruluğu kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
