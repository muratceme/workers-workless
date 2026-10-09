# Mevzuat Değişikliği Özeti · AI Agent

> Hukuk › Hukuk Müdürü · Workers / Workless

Resmî Gazete'nin günlük içindekiler sayfasındaki (fihrist) düzenlemeleri şirketin faaliyet alanına göre süzer.
İlgili düzenlemelerin tam metinlerinden **özet, değişiklikler, yapılacaklar ve yürürlük takvimi** çıkarır ve bir
mevzuat bülteni taslağı yazar. Her değişiklik metinden birebir alıntıyla desteklenir.

## Nasıl çalışır?

1. **Kod (fihrist):**
   - Fihristten Resmî Gazete tarihini, sayısını, bölümü, türü ve başlıkları okur. Bölümler Yasama, Yürütme ve
     İdare, Yargı ve İlân'dır; türler Kanun, Cumhurbaşkanı Kararı, Yönetmelik, Tebliğ gibi başlıklardır.
   - İlân bölümü varsayılan olarak alınmaz (`--ilanlar` ile alınır).
   - Profildeki anahtar kelimeleri başlıklarda arar.
2. **Model (sınıflandırma):**
   - Yalnız başlıklar ve şirket profili gönderilir.
   - Model her başlığa ilgi düzeyi (yüksek / orta / düşük / ilgisiz), konu, gerekçe ve etkilenen birimleri yazar.
   - Kod kimlikleri ve birim adlarını doğrular; profilde olmayan birim çıkarılır.
   - Anahtar kelime geçtiği hâlde "ilgisiz" denen başlık kontrol listesine düşer.
3. **Kod (yürürlük):** Tam metni verilen düzenlemede "yürürlüğe girer" içeren maddeyi **metinden birebir** çıkarır ve
   tarihleri hesaplar:
   - "1/1/2027 tarihinde" → 01.01.2027.
   - "yayımı tarihinde" → Resmî Gazete tarihi.
   - "yayımını izleyen ayın başında" → sonraki ayın 1'i.
   - Diğer ifadeler (ör. "yayımından altı ay sonra") için "metne bakın" yazılır.
4. **Model (özet):**
   - İlgisi yüksek veya orta olan ve tam metni bulunan düzenlemeler için özet, değişiklikler, yapılacaklar
     (birim ve zamanlama) ve belirsizlikleri yazar.
   - Her değişiklik için metinden birebir alıntı verir. Kod alıntının metinde geçtiğini doğrular; doğrulanamayan
     değişiklik rapora alınmaz.

`--model-yok` ile API kullanılmaz: yalnız fihrist, anahtar kelime taraması ve yürürlük takvimi çıkar.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                   # örnek: kurgusal fihrist + 2 tam metin
python agent.py --fihrist rg_2026-10-09.txt --metinler metinler/ --profil profil.txt
python agent.py --fihrist rg_2026-10-08.txt rg_2026-10-09.txt --profil profil.txt --model-yok
```

### Girdileri hazırlama

- **Fihrist:** resmigazete.gov.tr'de günün içindekiler sayfasını açın. Sayfayı seçip kopyalayın ve `.txt` olarak
  kaydedin.
  - Başlıklar `—` ile başlayan satırlardır. Bölüm ve tür başlıkları (YÖNETMELİKLER gibi) büyük harfle yazılır.
  - İlk satırlarda tarih bulunmalıdır, örneğin "Resmî Gazete 09.10.2026 Sayı: ...". Tarih dosya adında da olabilir:
    `rg_2026-10-09.txt`.
- **Tam metinler (isteğe bağlı):** İlgili görünen düzenlemenin metnini kopyalayıp `metinler/` klasörüne `.txt`
  olarak kaydedin. **İlk satır düzenlemenin başlığı** olmalıdır; başlık, fihristle kelime benzerliğine göre
  eşleştirilir. İlk çalıştırmada `--model-yok` ile hangi başlıkların ilgili olduğunu görüp yalnız onların metnini
  kaydedebilirsiniz.
- **Profil:** `Alan: değer` satırları. `Birimler:` ve `Anahtar kelimeler:` satırları `;` ile ayrılır (örnek:
  `ornek_veri/profil.txt`). Faaliyet, sektör, çalışan sayısı ve tesisleri yazmak sınıflandırmayı iyileştirir.

## Çıktı

- `mevzuat_ozeti.xlsx`:
  - `Özet`.
  - `İlgili Düzenlemeler`: ilgi, konu, gerekçe, birimler, özet, yürürlük; boş "Karar / Sorumlu" sütunu.
  - `Değişiklikler`: doğrulanmış alıntılarla.
  - `Yapılacaklar`: birim ve zamanlama; boş "Sorumlu" ve "Durum" sütunları.
  - `Yürürlük Takvimi`: tarih sırasıyla, metinden birebir yürürlük hükmü.
  - `Tüm Başlıklar`: anahtar kelime, ilgi, kontrol notları.
  - `Uyarılar`: yakında yürürlüğe girecek ilgili düzenleme (30 gün), tam metni gereken düzenleme, eşleşmeyen metin,
    doğrulanamayan alıntı.

  Mor hücreler yapay zekâ tarafından yazılmıştır.
- `mevzuat_ozeti_bulten.md`: paylaşılabilir bülten taslağı.

## Dikkat

- **Hukuki görüş değildir.** Sınıflandırma ve özetler taslaktır; hukuk biriminin metni okuyarak onaylaması gerekir.
  Özellikle "ilgisiz" ve "düşük" sınıflarını gözden geçirin.
- **Yalnız verilen metin:** Değişiklik düzenlemeleri genellikle ana metnin yalnız değişen kısmını içerir. Önceki
  hükmün ne olduğu metinde yoksa model bunu bilemez. Belirsizlikler sütununu okuyun.
- **Kaynak:** Paket internetten metin indirmez. Güncel ve resmî metin için Resmî Gazete ve mevzuat.gov.tr esastır.
- **Gönderilen veri:** Resmî Gazete metinleri kamuya açıktır. Profil dosyasına ticari sır veya kişisel veri yazmayın.
  Yerel model için `WW_PROVIDER=ollama` kullanabilirsiniz.
- **Örnek veri kurgusaldır:** Örnek fihrist, düzenleme metinleri ve şirket gerçek değildir. Gerçek Resmî Gazete
  içeriği değildir.

## Testler

Testler gerçek API çağırmaz; model yanıtları sahte fonksiyonla verilir. Şunlar test edilir:
- Fihrist ayrıştırma (iki satıra bölünmüş başlık, ilân bölümü) ve Türkçe başlık yazımı.
- Yürürlük tarihi kuralları; anahtar kelime taraması.
- Model yanıtlarının doğrulanması: geçersiz kimlik, profilde olmayan birim, metinde olmayan alıntı.
- İstem bölümlerinin ayrı gönderilmesi, Excel, bülten ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Hukuki görüş yerine
geçmez. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
