# Çağrı Kalite Değerlendirmesi · AI Agent

> Müşteri Hizmetleri ve Çağrı Merkezi › Müşteri Deneyimi Uzmanı · Workers / Workless

Görüşme dökümünü **kendi kalite formunuzdaki** kriterlere göre değerlendirir. Her kriter için sonuç, gerekçe ve
dökümden birebir alıntı verir; puanı hesaplar, kritik hataları işaretler ve temsilciye verilecek **geri bildirim
notunu** taslak olarak yazar. Alıntısı doğrulanamayan kriterler puana katılmaz ve insan incelemesine ayrılır.

## Nasıl çalışır?

1. **Kod (hazırlık):**
   - Dökümü okur: `Temsilci:` ve `Müşteri:` ile başlayan satırlar; baştaki zaman damgaları atılır. Dosyanın
     başındaki `Temsilci Adı:`, `Tarih:`, `Çağrı No:` satırları bilgi olarak alınır.
   - Telefon, e-posta, TCKN, IBAN ve kart numaralarını maskeler.
2. **Model (değerlendirme):** Her kriter için `karşılandı`, `kısmen`, `karşılanmadı` veya `uygulanamaz` sonucunu,
   gerekçeyi ve temsilcinin sözlerinden birebir bir alıntıyı verir. Güçlü yönleri, gelişim alanlarını ve geri
   bildirim notunu yazar.
3. **Kod (doğrulama ve puan):**

   | Durum | Sonuç |
   |---|---|
   | "Karşılandı" / "kısmen" ama alıntı temsilci sözlerinde birebir yok | İnceleme (insan) |
   | Alıntı yalnız müşterinin sözlerinde geçiyor | İnceleme (insan) |
   | Model "karşılandı" dedi ama formdaki anahtar ifadelerden hiçbiri temsilci sözlerinde yok | İnceleme (insan) |
   | Türü "İhlal" olan kriterde "karşılanmadı" ama ihlali gösteren alıntı yok | İnceleme (insan) |
   | Model kriter için sonuç döndürmedi | Değerlendirilmedi |
   | Formda olmayan kriter kodu | Atılır |

   - **Puan** = Σ ağırlık × (karşılandı 1; kısmen 0,5; karşılanmadı 0) / Σ ağırlık × 100. "Uygulanamaz", "inceleme" ve
     "değerlendirilmedi" kriterleri paya ve paydaya girmez.
   - **Kritik hata:** Kritik bir kriter karşılanmadıysa toplam puan 0 olur; ham puan ayrıca gösterilir.
     `--kritik-sifirlama-yok` ile puan sıfırlanmaz, yalnız işaretlenir.
   - Kritik kriter incelemeye düştüyse yüksek önemli uyarı verilir.

### Kalite formu

| Sütun | Açıklama |
|---|---|
| Kod, Bölüm, Kriter, Açıklama | Açıklama modelin kriteri doğru anlaması için önemlidir; beklenen davranışı yazın |
| Puan | Kriterin ağırlığı |
| Kritik | Evet / Hayır |
| Anahtar İfadeler | İsteğe bağlı; `|` ile ayrılmış. Kriter ancak bu ifadelerden biri temsilci sözlerinde geçiyorsa "karşılandı" kalır (ör. `talep numara|kayıt numara`) |
| Tür | `Davranış` (yapılması gereken; varsayılan) veya `İhlal` (yapılmaması gereken, ör. yasaklı ifade, gereksiz veri isteme). İhlal türünde "karşılandı" için alıntı aranmaz; "karşılanmadı" için ihlali gösteren alıntı aranır |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek: 2 görüşme, 10 kriterlik form
python agent.py --dokum gorusmeler/ --form kalite_formu.xlsx
python agent.py --dokum gorusme_0412.txt --form kalite_formu.csv --kritik-sifirlama-yok
```

Döküm biçimi:

```text
Temsilci Adı: Örnek Temsilci
Tarih: 05.10.2026
Temsilci: Örnek İletişim'e hoş geldiniz, size nasıl yardımcı olabilirim?
Müşteri: Faturam hakkında bilgi almak istiyorum.
```

## Çıktı

`cagri_kalite_degerlendirmesi.xlsx`:
- `Özet`: görüşme başına puan, ham puan, kritik hata, inceleme bekleyen kriter sayısı, temsilcinin konuşma payı; boş
  "Değerlendiren" sütunu.
- `Kriter Sonuçları`: sonuç, alınan puan, gerekçe, doğrulanmış alıntı, kontrol notu; boş "Uzman Kararı" sütunu.
- `Geri Bildirim`: güçlü yönler, gelişim alanları, taslak not.
- `Uyarılar`.

Mor hücreler yapay zekâ tarafından yazılmıştır.

## Dikkat

- **Sonuçlar taslaktır.** Puanlar performans değerlendirmesinde, primde veya disiplin sürecinde kullanılmadan önce
  bir kalite uzmanı kaydı dinleyerek onaylamalıdır. Çalışan hakkında yalnız otomatik değerlendirmeye dayanan karar
  verilmemelidir (KVKK md. 11).
- **Döküm sınırı:** Model yalnız yazıya dökülmüş sözleri görür. Ses tonu, bekleme süresi, sessizlik ve ekranda
  yapılan işlemler değerlendirilemez. Otomatik yazıya dökme hataları sonucu etkiler.
- **Kişisel veri:** Maskeleme yardımcıdır, garanti değildir; adlar ve adresler maskelenmez. Görüşme kayıtlarının
  işlenmesi ve API'ye gönderilmesi için aydınlatma ve yurt dışına aktarım şartlarını kontrol edin. Yerel model için
  `WW_PROVIDER=ollama` kullanabilirsiniz.
- **Örnek veri:** Örnek görüşmeler, firma ve kişiler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Form (ağırlık, kritik, ihlal türü) ve döküm okuma; maskeleme (telefon, kart).
- Alıntının temsilci / müşteri sözlerinde aranması.
- İnceleme kuralları: metinde olmayan alıntı, müşteri sözü, anahtar ifade eksikliği.
- Puan hesabı (elle), kritik hata ve sıfırlama seçeneği, formda olmayan kod, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
